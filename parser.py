#!/usr/bin/env python3
import re, json, sys, argparse
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def classify_params(body, params):
    result = {}
    for p in params:
        esc = re.escape(p)
        ptr_write   = re.search(rf"[\*\(]\s*{esc}\s*\)?\s*(?:[+\-*/]?=|[+\-]{{2}})", body)
        arrow_write = re.search(rf"\b{esc}\s*->\s*\w+\s*=", body)
        inc_dec     = re.search(rf"(?:\+\+|--)\s*{esc}\b|\b{esc}\s*(?:\+\+|--)", body)
        write_api   = re.search(rf"\b\w*(?:Write|Set)\w*\s*\([^;]*\b{esc}\b", body)

        is_written = bool(ptr_write or arrow_write or inc_dec or write_api)
        is_read    = bool(re.search(rf"\b{esc}\b", body))

        result[p] = "INOUT" if (is_written and is_read) else ("OUT" if is_written else "IN")
    return result

def get_trigger_comment(comments, pos):
    """Return last comment block before pos containing a '- triggered on|by …' line."""
    trig_rx = re.compile(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", re.IGNORECASE)
    best_end, best_txt = -1, ""
    for end, txt in comments:
        if end <= pos and end > best_end and trig_rx.search(txt):
            best_end, best_txt = end, txt
    return best_txt

def parse_file(src):
    rows = []
    reserved = {"if","for","while","switch","do","else","case","sizeof","abs","return"}
    exclude_invoked = {
        "VStdLib_MemCpy","VStdLib_MemSet","VStdLib_MemCmp",
        "memcmp","memcpy","memset","sizeof","abs","return"
    }

    # 1) collect comment blocks
    comments = [(m.end(), m.group(0)) for m in re.finditer(r"/\*[\s\S]*?\*/", src)]

    # 2a) AUTOSAR runnables: FUNC(return,code) name(
    runnable_rx = re.compile(
        r'^[ \t]*FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)\s*'
        r'([A-Za-z_]\w*)\s*\(',
        re.MULTILINE
    )

    # 2b) static helpers
    static_rx = re.compile(r'''
        ^[ \t]*static\s+(?:inline\s+)?                  # static or static inline
        (?:
          FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)  # FUNC macro → group1
        |
          ([A-Za-z_]\w*)                                # direct returnType → group2
        )
        \s+([A-Za-z_]\w*)\s*\(                          # name → group3
        ''', re.MULTILINE | re.VERBOSE)

    # 2c) global helpers
    global_rx = re.compile(r'''
        ^[ \t]*(?!static\b)(?!FUNC\b)                   # no static, no FUNC
        ([A-Za-z_]\w*(?:\s*\*+)?)\s+                    # returnType → group1
        (?!(?:if|for|while|switch|do|else|case)\b)      # skip reserved
        ([A-Za-z_]\w*)\s*\(                             # name → group2
        ''', re.MULTILINE | re.VERBOSE)

    # signature regex for Syntax field
    sig_rx = re.compile(r'''
        ^[ \t]*                                        # indent
        (?:static\s+)?(?:inline\s+)?                   # optional
        (?:FUNC\([^)]*\)\s*)?                          # optional FUNC()
        (?P<ret>[\w\*\s]+?)\s+                         # return type
        (?P<name>[A-Za-z_]\w*)\s*                      # name
        \((?P<params>[^)]*)\)                          # params
        ''', re.MULTILINE | re.VERBOSE)

    def extract(m, fnType):
        # pick return type & name
        if fnType == "Static":
            retType = m.group(1) or m.group(2)
            name    = m.group(3)
        else:
            retType, name = m.group(1), m.group(2)

        # find raw_params by counting parentheses
        L = len(src)
        depth, i = 1, m.end()
        while i < L and depth:
            depth += src[i] == "("
            depth -= src[i] == ")"
            i += 1
        raw_params = src[m.end(): i-1].strip()

        # build syntax
        snippet = src[m.start():]
        sig_m = sig_rx.match(snippet)
        if sig_m:
            syntax = f"{sig_m.group('ret').strip()} {sig_m.group('name')}({sig_m.group('params').strip()})"
        else:
            syntax = f"{retType} {name}({raw_params})"

        # skip if no body
        pos = i
        while pos < L and (src[pos].isspace() or src.startswith("/*", pos)):
            pos = src.find("*/", pos)+2 if src.startswith("/*", pos) else pos+1
        if pos >= L or src[pos] != "{":
            return

        # extract body + strip comments
        brace_idx, depth, j = pos, 1, pos+1
        while j < L and depth:
            depth += src[j] == "{"
            depth -= src[j] == "}"
            j+=1
        body = src[brace_idx+1: j-1]
        code = re.sub(r'/\*[\s\S]*?\*/|//.*', '', body)

        # triggers
        if fnType == "Runnable":
            cm = get_trigger_comment(comments, m.start())
            trigs = re.findall(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", cm, re.IGNORECASE)
            trigger = "; ".join(t.strip() for t in trigs)
        else:
            trigger = ""

        # params → names
        parts = [p.strip() for p in re.split(r",(?![^(]*\))", raw_params) if p.strip()]
        names = []
        for p in parts:
            mm = re.search(r"\b([A-Za-z_]\w*)\s*$", p)
            names.append(mm.group(1) if mm else p)

        # IN/OUT
        dirs = classify_params(body, names)
        for orig,nm in zip(parts,names):
            if orig.startswith(("P2CONST(","P2VAR(")):
                dirs[nm] = "OUT"
        inP  = [p for p in names if dirs[p] in ("IN","INOUT")]
        outP = [p for p in names if dirs[p] in ("OUT","INOUT")]

        # RTE APIs
        inputs  = sorted({x.split("(")[0] for x in re.findall(
                     r"\bRte_(?:Read|DRead|IRead|Receive|IReadRef|IrvRead|IsUpdated|Mode_)[\w_]*\s*\(",
                     body)})
        outputs = sorted({x.split("(")[0] for x in re.findall(
                     r"\bRte_(?:Write|IrvWrite|IWrite|IWriteRef|Switch)[\w_]*\s*\(",
                     body)})

        # invoked
        calls = {x.split("(")[0] for x in re.findall(r"\bRte_Call_[\w_]+\s*\(", code)}
        plain = re.findall(r"\b([A-Za-z_]\w*)\s*\(", code)
        locals_ = {c for c in plain
                   if c not in reserved
                   and c not in exclude_invoked
                   and not c.startswith("Rte_")
                   and c != name}
        invoked = sorted(c for c in calls|locals_
                         if not re.fullmatch(r"[A-Z][A-Z0-9_]*", c))

        # used types
        used = sorted({t for t in re.findall(
                      r"\b([A-Za-z_]\w*)\s+[A-Za-z_]\w*\s*(?:[=;])", body)
                      if t.lower() not in reserved})
        
        sync_async = ""
        reentrancy = ""

        rows.append({
            "name":      name,
            "syntax":    syntax,
            "ret":       retType,
            "inParams":  inP,
            "outParams": outP,
            "fnType":    fnType,
            "trigger":   trigger,
            "inputs":    inputs,
            "outputs":   outputs,
            "invoked":   invoked,
            "used":      used,
            "Sync_Async": sync_async,
            "Reentrancy": reentrancy
        })

    # run extractors
    for m in runnable_rx.finditer(src):
        extract(m, "Runnable")
    for m in static_rx.finditer(src):
        extract(m, "Static")
    for m in global_rx.finditer(src):
        extract(m, "Global")

    return rows

def write_docx(source_path, swcName, rows):
    """Generate a Word .docx per the AUTOSAR‐API template alongside the Excel."""
    def shade_cell(cell, rgb="9D9D9D"):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), rgb)
        tcPr.append(shd)

    doc = Document()
    for r in rows:
        doc.add_paragraph(f"[{r['name']}]", style='Heading 1')
        tbl = doc.add_table(rows=0, cols=2)
        tbl.style = 'Table Grid'

        def add_row(label, value):
            c0, c1 = tbl.add_row().cells
            c0.text = label
            c1.text = value or ""
            shade_cell(c0)

        add_row("Service Name",        r['name'])
        add_row("Syntax",              r['syntax'])
        add_row("Sync/Async",          "")
        add_row("Reentrancy",          "")
        add_row("Parameters (in)",     ", ".join(r['inParams']))
        add_row("Parameters (inout)",  "")
        add_row("Parameters (out)",    ", ".join(r['outParams']))
        add_row("Return value",        r['ret'])
        add_row("Description",         "")
        add_row("Triggers",            r['trigger'])
        add_row("Inputs",              ", ".join(r['inputs']))
        add_row("Outputs",             ", ".join(r['outputs']))
        add_row("Invoked Operations",  ", ".join(r['invoked']))
        add_row("Used Data Types",     ", ".join(r['used']))
        add_row("Available via",       "")
        doc.add_page_break()

    # save alongside the Excel (.xlsx) in the same directory
    excel_path = Path(source_path).with_suffix('.xlsx')
    docx_path  = excel_path.with_suffix('.docx')
    doc.save(docx_path)

    # print is shitty here don't use it -> will ruin the parsing of the output JSON file.
    # print(f"↪️  Written Word doc: {docx_path}", file=sys.stderr)

def main():
    ap = argparse.ArgumentParser(description="Extract and document Runnables")
    ap.add_argument("file", help="C source file path")
    args = ap.parse_args()

    try:
        src = Path(args.file).read_text(encoding="utf-8")
    except Exception as e:
        print(f"❌ Error reading {args.file}: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        rows = parse_file(src)
    except Exception as e:
        print(f"❌ Error parsing {args.file}: {e}", file=sys.stderr)
        sys.exit(1)

    # output JSON for compatibility
    print(json.dumps(rows, indent=2))

    # write the Word doc next to the Excel
    swcName = Path(args.file).stem
    write_docx(args.file, swcName, rows)

if __name__ == "__main__":
    main()
