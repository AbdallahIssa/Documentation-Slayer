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
    """Pick the last comment block before pos that contains a trigger line."""
    trig_rx = re.compile(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", re.IGNORECASE)
    best_end, best_txt = -1, ""
    for end, txt in comments:
        if end <= pos and end > best_end and trig_rx.search(txt):
            best_end, best_txt = end, txt
    return best_txt

def parse_file(src):
    rows = []
    reserved = {"if","for","while","switch","do","else","case"}
    exclude_invoked = {"VStdLib_MemCpy","VStdLib_MemSet","memcpy","memset","sizeof"}

    # 1) gather all /* … */ blocks for trigger lookup
    comments = [(m.end(), m.group(0)) for m in re.finditer(r"/\*[\s\S]*?\*/", src)]

    # 2) regexes that go only as far as the first '('
    runnable_rx = re.compile(
        r'^[ \t]*FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)\s*'  # capture return
        r'([A-Za-z_]\w*)\s*\('                                           # capture name + '('
        , re.MULTILINE
    )
    static_rx = re.compile(
        r'^[ \t]*static\s+([A-Za-z_]\w*)\s+'   # capture return type
        r'([A-Za-z_]\w*)\s*\('                 # capture name + '('
        , re.MULTILINE
    )

    def extract(m, fnType):
        retType, name = m.group(1), m.group(2)
        params_start  = m.end()                     # position right _after_ '('
        L = len(src)

        # 3) manual parenthesis-count to find matching ')'
        depth, i = 1, params_start
        while i < L and depth:
            if   src[i] == "(": depth += 1
            elif src[i] == ")": depth -= 1
            i += 1
        if depth != 0:
            return  # unbalanced—skip

        raw_params = src[params_start : i-1].strip()

        # 4) now skip whitespace/comments to find the '{' of the body
        pos = i
        while pos < L:
            if src.startswith("/*", pos):
                pos = src.find("*/", pos) + 2
            elif src[pos].isspace():
                pos += 1
            else:
                break
        if pos >= L or src[pos] != "{":
            return  # this was a prototype, not a definition

        # 5) brace-count the body
        brace_idx, depth, j = pos, 1, pos+1
        while j < L and depth:
            if   src[j] == "{": depth += 1
            elif src[j] == "}": depth -= 1
            j += 1
        body = src[brace_idx+1 : j-1]

        # 6) extract the trigger (runnables only)
        if fnType == "Runnable":
            cm      = get_trigger_comment(comments, m.start())
            tm      = re.search(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", cm, re.IGNORECASE)
            trigger = tm.group(1).strip() if tm else ""
        else:
            trigger = ""

        # 7) split params (handles nested commas) → names
        parts = [p.strip() for p in re.split(r",(?![^(]*\))", raw_params) if p.strip()]
        names = []
        for p in parts:
            mm = re.search(r"\b([A-Za-z_]\w*)\s*$", p)
            names.append(mm.group(1) if mm else p)

        # 8) classify IN/OUT/INOUT
        dirs = classify_params(body, names)
        #  → force any P2CONST(...) param to OUT
        for original, nm in zip(parts, names):
            if original.startswith("P2CONST("):
                dirs[nm] = "OUT"

        inP  = [p for p in names if dirs[p] in ("IN","INOUT")]
        outP = [p for p in names if dirs[p] in ("OUT","INOUT")]

        # 9) RTE APIs
        inputs  = sorted({x.split("(")[0] for x in re.findall(r"\bRte_Read_[\w_]+\s*\(", body)})
        outputs = sorted({x.split("(")[0] for x in re.findall(r"\bRte_(?:Write|IrvWrite)_[\w_]+\s*\(", body)})
        calls   = sorted({x.split("(")[0] for x in re.findall(r"\bRte_Call_[\w_]+\s*\(", body)})

        # 10) other local calls, minus reserved & excludes
        plain   = re.findall(r"\b([A-Za-z_]\w*)\s*\(", body)
        locals_ = sorted({
            c for c in plain
            if c not in reserved
            and c not in exclude_invoked
            and not c.startswith("Rte_")
            and c != name
        })
        invoked = sorted(set(calls + locals_))

        # 11) used data types
        used = sorted({
            t for t in re.findall(r"\b([A-Za-z_]\w*)\s+[A-Za-z_]\w*\s*(?:[=;])", body)
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

    # 12) scan for definitions only
    for m in runnable_rx.finditer(src):
        extract(m, "Runnable")
    for m in static_rx.finditer(src):
        extract(m, "Static")

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
        print(f"❌ Error parsing {args.file}: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
