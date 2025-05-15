#!/usr/bin/env python3
import re, json, sys, argparse

def classify_params(body, params):
    """Heuristic IN/OUT/INOUT detection."""
    result = {}
    for p in params:
        esc = re.escape(p)
        ptr_write   = re.search(rf"[\*\(]\s*{esc}\s*\)?\s*(?:[+\-*/]?=|[+\-]{{2}})", body)
        arrow_write = re.search(rf"\b{esc}\s*->\s*\w+\s*=", body)
        inc_dec     = re.search(rf"(?:\+\+|--)\s*{esc}\b|\b{esc}\s*(?:\+\+|--)", body)
        write_api   = re.search(rf"\b\w*(?:Write|Set)\w*\s*\([^;]*\b{esc}\b", body)

        is_written  = bool(ptr_write or arrow_write or inc_dec or write_api)
        is_read     = bool(re.search(rf"\b{esc}\b", body))

        if is_written:
            result[p] = "INOUT" if is_read else "OUT"
        else:
            result[p] = "IN"
    return result

def get_trigger_comment(comments, pos):
    """
    Return the last comment block before pos that actually has
    one or more '- triggered on|by …' lines.
    """
    trig_rx = re.compile(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", re.IGNORECASE)
    best_end, best_txt = -1, ""
    for end, txt in comments:
        if end <= pos and end > best_end and trig_rx.search(txt):
            best_end, best_txt = end, txt
    return best_txt

def parse_file(src):
    rows = []
    reserved = {"if","for","while","switch","do","else","case"}
    exclude_invoked = {"VStdLib_MemCpy","VStdLib_MemSet","memcpy","memset","sizeof", "abs"}

    # 1) collect comment blocks for later trigger lookup
    comments = [(m.end(), m.group(0)) for m in re.finditer(r"/\*[\s\S]*?\*/", src)]

    # 2a) AUTOSAR runnables: FUNC(return,code) name(
    runnable_rx = re.compile(
        r'^[ \t]*FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)\s*'  # return type
        r'([A-Za-z_]\w*)\s*\('                                           # name + '('
        , re.MULTILINE
    )
    # 2b) static helpers: static returnType name(
    static_rx = re.compile(
        r'^[ \t]*static\s+([A-Za-z_]\w*)\s+'  # returnType
        r'([A-Za-z_]\w*)\s*\('                # name + '('
        , re.MULTILINE
    )
    # 2c) global helpers:  returnType  name(   …but NOT static, NOT FUNC(...), NOT reserved words
    global_rx = re.compile(
        r'''
        ^[ \t]*                        # line start + optional indent
        (?!static\b)                   # not a static function
        (?!FUNC\b)                     # not the FUNC(...) macro
        ([A-Za-z_]\w*(?:\s*\*+)?)\s+   # return type (group 1)
        (?!(?:if|for|while|switch|do|else|case)\b)  # EXCLUDE reserved keywords
        ([A-Za-z_]\w*)\s*\(            # function name (group 2) + '('
        ''',
        re.MULTILINE | re.VERBOSE
    )

    def extract(m, fnType):
        retType, name = m.group(1), m.group(2)
        L = len(src)

        # 3) find matching ')' by counting
        depth, i = 1, m.end()
        while i < L and depth:
            if   src[i] == "(": depth += 1
            elif src[i] == ")": depth -= 1
            i += 1
        raw_params = src[m.end(): i-1].strip()

        # 4) skip prototypes: next non-space/comment must be '{'
        pos = i
        while pos < L:
            if src.startswith("/*", pos):
                pos = src.find("*/", pos)+2
            elif src[pos].isspace():
                pos += 1
            else:
                break
        if pos >= L or src[pos] != "{":
            return  # not a definition

        # 5) brace-count to extract body
        brace_idx, depth, j = pos, 1, pos+1
        while j < L and depth:
            if   src[j] == "{": depth += 1
            elif src[j] == "}": depth -= 1
            j += 1
        body = src[brace_idx+1 : j-1]
        # ─── strip out ALL comments (/* … */ and // … ) so invokedOps sees only real code
        code = re.sub(r'/\*[\s\S]*?\*/|//.*', '', body)

        # 6) MULTI-LINE triggers for runnables
        if fnType == "Runnable":
            cm = get_trigger_comment(comments, m.start())
            all_trigs = re.findall(
                r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)",
                cm,
                re.IGNORECASE
            )
            trigger = "; ".join(t.strip() for t in all_trigs)
        else:
            trigger = ""

        # 7) split params → names
        parts = [p.strip() for p in re.split(r",(?![^(]*\))", raw_params) if p.strip()]
        names = []
        for p in parts:
            mm = re.search(r"\b([A-Za-z_]\w*)\s*$", p)
            names.append(mm.group(1) if mm else p)

        # 8) IN/OUT/INOUT + force P2CONST(...) and P2VAR(...) → OUT
        dirs = classify_params(body, names)
        for orig, nm in zip(parts, names):
            if (orig.startswith("P2CONST(") or orig.startswith("P2VAR(")):
                dirs[nm] = "OUT"

        inP  = [p for p in names if dirs[p] in ("IN","INOUT")]
        outP = [p for p in names if dirs[p] in ("OUT","INOUT")]

        # 9) RTE APIs
        # match any of: Read, DRead, IRead, Receive, IReadRef, IrvRead, IsUpdated, Mode_* 
        inputs = sorted({
            x.split("(")[0]
            for x in re.findall(
                r"\bRte_(?:Read|DRead|IRead|Receive|IReadRef|IrvRead|IsUpdated|CData|Mode_)[\w_]*\s*\(",
                body
            )
        })

        # match any of: Write, IrvWrite, IWrite, IWriteRef, Switch* 
        outputs = sorted({
            x.split("(")[0]
            for x in re.findall(
                r"\bRte_(?:Write|IrvWrite|IWrite|IWriteRef|Switch)[\w_]*\s*\(",
                body
            )
        })

        calls = sorted({x.split("(")[0] for x in re.findall(r"\bRte_Call_[\w_]+\s*\(",    code)})

        # 10) other locals minus reserved/excluded
        plain = re.findall(r"\b([A-Za-z_]\w*)\s*\(", code)
        locals_ = sorted({
            c for c in plain
            if c not in reserved
            and c not in exclude_invoked
            and not c.startswith("Rte_")
            and c != name
        })
        invoked = sorted(set(calls + locals_))

        # 11) used data types, excluding "return"
        used = sorted({
            t for t in re.findall(
                r"\b([A-Za-z_]\w*)\s+[A-Za-z_]\w*\s*(?:[=;])",
                body
            ) if t.lower() != "return"
        })

        rows.append({
            "name":      name,
            "ret":       retType,
            "inParams":  inP,
            "outParams": outP,
            "fnType":    fnType,
            "trigger":   trigger,
            "inputs":    inputs,
            "outputs":   outputs,
            "invoked":   invoked,
            "used":      used
        })

    # 12) run extraction
    for m in runnable_rx.finditer(src):
        extract(m, "Runnable")
    for m in static_rx.finditer(src):
        extract(m, "Static")
    for m in global_rx.finditer(src):
        extract(m, "Global")

    return rows

def main():
    ap = argparse.ArgumentParser(description="Extract runnables/static info")
    ap.add_argument("file", help="C source file path")
    args = ap.parse_args()

    try:
        src = open(args.file, encoding="utf-8").read()
    except Exception as e:
        print(f"❌ Error reading {args.file}: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        data = parse_file(src)
        print(json.dumps(data, indent=2))
    except Exception as e:
        print(f"❌ Error parsing {args.file}: {e}", file=sys.stderr) # only for Debugggggggggggggggggggggggg
        sys.exit(1)

if __name__ == "__main__":
    main()
