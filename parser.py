#!/usr/bin/env python3
import re
import json
import sys
import argparse
import os
import platform
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tkinter as tk
from tkinter import ttk, filedialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

def open_file(path: str):
    """Open a file with the default OS application."""
    system = platform.system()
    if system == "Windows":
        os.startfile(path)
    elif system == "Darwin":
        os.system(f'open "{path}"')
    else:
        os.system(f'xdg-open "{path}"')

def write_excel(file_path: str, rows: list[dict], sel_fields: list[str]):
    """
    Write an .xlsx with one sheet named
    "Runnables and static functions".
    Only the sel_fields columns are written.
    Headers are styled bold+red font on yellow fill.
    """
    ALL_HEADERS = [
      'Name', 'Syntax', 'Return Value', 'In-Parameters', 'Out-Parameters',
      'Function Type', 'Description', 'Sync/Async', 'Reentrancy',
      'Triggers', 'Inputs', 'Outputs',
      'Invoked Operations', 'Used Data Types'
    ]
    headers = [h for h in ALL_HEADERS if h in sel_fields]

    path = Path(file_path)
    if path.exists():
        wb = load_workbook(str(path))
    else:
        wb = Workbook()

    sheet_name = "Runnables and static functions"
    # remove all other sheets
    for name in list(wb.sheetnames):
        if name != sheet_name:
            wb.remove(wb[name])
    # get or create our sheet
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)

    # write header
    ws.append(headers)
    yellow = PatternFill(fill_type="solid", fgColor="FFFFFF00")
    red    = Font(bold=True, color="FFFF0000")
    for cell in ws[1]:
        cell.fill = yellow
        cell.font = red

    # append rows
    for r in rows:
        row = []
        for h in headers:
            if h == 'Name':
                row.append(r['name'])
            elif h == 'Syntax':
                row.append(r['syntax'])
            elif h == 'Return Value':
                row.append(r['ret'])
            elif h == 'In-Parameters':
                row.append(", ".join(r['inParams']))
            elif h == 'Out-Parameters':
                row.append(", ".join(r['outParams']))
            elif h == 'Function Type':
                row.append(r['fnType'])
            elif h == 'Description':
                row.append(r.get('description', ''))
            elif h == 'Sync/Async':
                row.append(r['Sync_Async'])
            elif h == 'Reentrancy':
                row.append(r['Reentrancy'])
            elif h == 'Triggers':
                row.append(r['trigger'])
            elif h == 'Inputs':
                row.append(", ".join(r['inputs']))
            elif h == 'Outputs':
                row.append(", ".join(r['outputs']))
            elif h == 'Invoked Operations':
                row.append(", ".join(r['invoked']))
            elif h == 'Used Data Types':
                row.append(", ".join(r['used']))
        ws.append(row)

    # auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            val = str(cell.value or "")
            max_length = max(max_length, len(val))
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(str(path))

def write_markdown(file_path: str, rows: list[dict], sel_fields: list[str]):
    """
    Write a Markdown file with a table for each function.
    Only the sel_fields are included in each table.
    """
    FIELD_GETTERS = {
      'Name':             lambda r: r['name'],
      'Syntax':           lambda r: f"`{r['syntax']}`",
      'Sync/Async':       lambda r: f"`{r['Sync_Async']}`",
      'Reentrancy':       lambda r: f"`{r['Reentrancy']}`",
      'Return Value':     lambda r: f"`{r['ret']}`",
      'In-Parameters':    lambda r: ", ".join(r['inParams']),
      'Out-Parameters':   lambda r: ", ".join(r['outParams']),
      'Function Type':    lambda r: r['fnType'],
      'Description':      lambda r: r.get('description', ''),
      'Triggers':         lambda r: r['trigger'],
      'Inputs':           lambda r: ", ".join(r['inputs']),
      'Outputs':          lambda r: ", ".join(r['outputs']),
      'Invoked Operations': lambda r: ", ".join(r['invoked']),
      'Used Data Types':    lambda r: ", ".join(r['used']),
    }

    lines = []
    # 1) record which fields were selected
    lines.append(f"**Selected fields:** {', '.join(sel_fields)}")
    lines.append("")

    # 2) one table per function
    for r in rows:
        lines.append(f"## {r['name']}")
        lines.append("")
        lines.append("| Field | Value |")
        lines.append("|-------|-------|")
        for label in sel_fields:
            getter = FIELD_GETTERS.get(label)
            if getter:
                value = getter(r)
                lines.append(f"| {label} | {value} |")
        lines.append("")  # blank line between functions

    # 3) write to disk
    Path(file_path).write_text("\n".join(lines), encoding="utf-8")


def write_docx(source_path: str, swcName: str, rows: list[dict], sel_fields: list[str]):
    """
    Generate a Word .docx per AUTOSAR template.
    Only includes sel_fields.
    """
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

        if 'Name' in sel_fields:
            add_row("Service Name", r['name'])
        if 'Syntax' in sel_fields:
            add_row("Syntax", r['syntax'])
        if 'Sync/Async' in sel_fields:
            add_row("Sync/Async", r['Sync_Async'])
        if 'Reentrancy' in sel_fields:
            add_row("Reentrancy", r['Reentrancy'])
        if 'In-Parameters' in sel_fields:
            add_row("Parameters (in)", ", ".join(r['inParams']))
        if 'Out-Parameters' in sel_fields:
            add_row("Parameters (out)", ", ".join(r['outParams']))
        if 'Function Type' in sel_fields:
            add_row("Function Type", r['fnType'])
        if 'Description' in sel_fields:
            add_row("Description", r.get('description', ''))
        if 'Triggers' in sel_fields:
            add_row("Triggers", r['trigger'])
        if 'Inputs' in sel_fields:
            add_row("Inputs", ", ".join(r['inputs']))
        if 'Outputs' in sel_fields:
            add_row("Outputs", ", ".join(r['outputs']))
        if 'Invoked Operations' in sel_fields:
            add_row("Invoked Operations", ", ".join(r['invoked']))
        if 'Used Data Types' in sel_fields:
            add_row("Used Data Types", ", ".join(r['used']))

        doc.add_page_break()

    excel_path = Path(source_path).with_suffix('.xlsx')
    docx_path  = excel_path.with_suffix('.docx')
    doc.save(str(docx_path))

def classify_params(body: str, params: list[str]) -> dict:
    """Heuristic IN/OUT/INOUT detection."""
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

def get_trigger_comment(comments: list, pos: int) -> str:
    trig_rx = re.compile(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", re.IGNORECASE)
    best_end, best_txt = -1, ""
    for end, txt in comments:
        if end <= pos and end > best_end and trig_rx.search(txt):
            best_end, best_txt = end, txt
    return best_txt

def parse_file(src: str) -> list:
    rows = []
    reserved = {"if","for","while","switch","do","else","case","sizeof","abs","return"}
    exclude_invoked = {
        "VStdLib_MemCpy","VStdLib_MemSet","VStdLib_MemCmp",
        "memcmp","memcpy","memset","sizeof","abs","return"
    }

    comments = [(m.end(), m.group(0)) for m in re.finditer(r"/\*[\s\S]*?\*/", src)]

    runnable_rx = re.compile(
        r'^[ \t]*FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)\s*([A-Za-z_]\w*)\s*\(',
        re.MULTILINE
    )
    static_rx = re.compile(r'''
        ^[ \t]*static\s+(?:inline\s+)?                  
        (?:
          FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)  
        |
          ([A-Za-z_]\w*)                                
        )
        \s+([A-Za-z_]\w*)\s*\(                          
        ''', re.MULTILINE|re.VERBOSE)
    global_rx = re.compile(r'''
        ^[ \t]*(?!static\b)(?!FUNC\b)                   
        ([A-Za-z_]\w*(?:\s*\*+)?)\s+                    
        (?!(?:if|for|while|switch|do|else|case)\b)      
        ([A-Za-z_]\w*)\s*\(                             
        ''', re.MULTILINE|re.VERBOSE)
    sig_rx = re.compile(r'''
        ^[ \t]*                                        
        (?:static\s+)?(?:inline\s+)?                   
        (?:FUNC\([^)]*\)\s*)?                          
        (?P<ret>[\w\*\s]+?)\s+                         
        (?P<name>[A-Za-z_]\w*)\s*                      
        \((?P<params>[^)]*)\)                          
        ''', re.MULTILINE|re.VERBOSE)

    def extract(m, fnType):
        if fnType == "Static":
            retType = m.group(1) or m.group(2)
            name    = m.group(3)
        else:
            retType, name = m.group(1), m.group(2)

        # parameters
        L = len(src)
        depth, i = 1, m.end()
        while i < L and depth:
            depth += src[i] == "("
            depth -= src[i] == ")"
            i += 1
        raw_params = src[m.end():i-1].strip()

        # syntax
        snippet = src[m.start():]
        sig_m = sig_rx.match(snippet)
        if sig_m:
            syntax = f"{sig_m.group('ret').strip()} {sig_m.group('name')}({sig_m.group('params').strip()})"
        else:
            syntax = f"{retType} {name}({raw_params})"

        # skip prototypes
        pos = i
        while pos < L and (src[pos].isspace() or src.startswith("/*", pos)):
            pos = src.find("*/", pos) + 2 if src.startswith("/*", pos) else pos+1
        if pos >= L or src[pos] != "{":
            return

        # body
        brace_idx, depth, j = pos, 1, pos+1
        while j < L and depth:
            depth += src[j] == "{"
            depth -= src[j] == "}"
            j += 1
        body = src[brace_idx+1:j-1]
        code = re.sub(r'/\*[\s\S]*?\*/|//.*', '', body)

        # trigger
        if fnType == "Runnable":
            cm = get_trigger_comment(comments, m.start())
            trigs = re.findall(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", cm, re.IGNORECASE)
            trigger = "; ".join(t.strip() for t in trigs)
        else:
            trigger = ""

        # params names
        parts = [p.strip() for p in re.split(r",(?![^(]*\))", raw_params) if p.strip()]
        names = []
        for p in parts:
            mm = re.search(r"\b([A-Za-z_]\w*)\s*$", p)
            names.append(mm.group(1) if mm else p)

        # IN/OUT
        dirs = classify_params(body, names)
        for orig, nm in zip(parts, names):
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
        locals_ = {
            c for c in plain
            if c not in reserved
            and c not in exclude_invoked
            and not c.startswith("Rte_")
            and c != name
        }
        invoked = sorted(c for c in calls|locals_ if not re.fullmatch(r"[A-Z][A-Z0-9_]*", c))

        # used types
        used = sorted({
            t for t in re.findall(r"\b([A-Za-z_]\w*)\s+[A-Za-z_]\w*\s*(?:[=;])", body)
            if t.lower() not in reserved
        })

        # placeholders for GUI fields
        sync_async = ""
        reentrancy = ""
        description = ""

        rows.append({
            "name":       name,
            "syntax":     syntax,
            "ret":        retType,
            "inParams":   inP,
            "outParams":  outP,
            "fnType":     fnType,
            "trigger":    trigger,
            "inputs":     inputs,
            "outputs":    outputs,
            "invoked":    invoked,
            "used":       used,
            "Sync_Async": sync_async,
            "Reentrancy": reentrancy,
            "description": description
        })

    for m in runnable_rx.finditer(src):
        extract(m, "Runnable")
    for m in static_rx.finditer(src):
        extract(m, "Static")
    for m in global_rx.finditer(src):
        extract(m, "Global")

    return rows

def show_gui():
    fields = [
      "Name", "Description", "Syntax", "Triggers", "In-Parameters", "Out-Parameters",
      "Return Value", "Function Type", "Inputs", "Outputs",
      "Invoked Operations", "Used Data Types", "Sync/Async", "Reentrancy"
    ]
    formats = ["Excel", "Word", "MD"]

    root = tk.Tk()
    root.title("Documentation Slayer")
    root.configure(bg="#ececec")
    root.iconbitmap("vehiclevo_logo_Basic.ico")
    root.geometry("700x350")
    root.minsize(700, 350)
    root.maxsize(700, 350)
    root.minsize(700, 350)
    root.resizable(False, False) # no resizing

    root.update_idletasks()         # ensure winfo_width/height are accurate
    w = root.winfo_width()
    h = root.winfo_height()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

    field_vars  = {f: tk.BooleanVar(value=True) for f in fields}
    format_vars = {f: tk.BooleanVar(value=True) for f in formats}
    save_dir    = tk.StringVar(value=str(Path.cwd()))

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Select fields:").grid(row=0, column=0, sticky="w")
    for i, f in enumerate(fields, 1):
        chk = ttk.Checkbutton(frm, text=f, variable=field_vars[f])
        chk.grid(row=(i-1)//3+1, column=(i-1)%3, sticky="w", padx=15, pady=5)

    ttk.Label(frm, text="Select formats:").grid(row=6, column=0, sticky="w", pady=(10,0))
    for i, fmt in enumerate(formats, 7):
        chk = ttk.Checkbutton(frm, text=fmt, variable=format_vars[fmt])
        chk.grid(row=6 + (i-7)//3+1, column=(i-7)%3, sticky="w", padx=15, pady=5)

    def choose_dir():
        d = filedialog.askdirectory()
        if d:
            save_dir.set(d)
    ttk.Button(frm, text="Save to…", command=choose_dir).grid(row=10, column=0, sticky="w", pady=(10,0))
    ttk.Label(frm, textvariable=save_dir).grid(row=10, column=1, columnspan=2, sticky="w")

    def on_run():
        # Gather user selections
        sel_fields  = [f for f, v in field_vars.items() if v.get()]
        sel_formats = [f for f, v in format_vars.items() if v.get()]
        outdir      = Path(save_dir.get())

        # Ask for the C source file
        cfile = filedialog.askopenfilename(
            title="Select C source file",
            filetypes=[("C files","*.c"), ("All files","*.*")]
        )
        if not cfile:
            return

        # Parse the file
        src  = open(cfile, encoding="utf-8").read()
        rows = parse_file(src)
        stem = Path(cfile).stem

        # Precompute the Excel path (used for .xlsx and .docx output)
        xlsx_path = outdir / f"{stem}.xlsx"

        # Excel export
        if "Excel" in sel_formats:
            write_excel(str(xlsx_path), rows, sel_fields)
            open_file(str(xlsx_path))

        # Word export
        if "Word" in sel_formats:
            # write_docx takes the Excel path to derive the .docx alongside it
            write_docx(str(xlsx_path), stem, rows, sel_fields)
            open_file(str(xlsx_path.with_suffix('.docx')))

        # Markdown export
        if "MD" in sel_formats:
            md_path = outdir / f"{stem}.md"
            write_markdown(str(md_path), rows, sel_fields)
            open_file(str(md_path))

        # Close the GUI
        root.destroy()

    ttk.Button(frm, text="Run", command=on_run).grid(row=11, column=0, pady=15)
    root.mainloop()

def main():
    ap = argparse.ArgumentParser(description="Extract and document Runnables")
    ap.add_argument("file", nargs="?", default=None, help="C source file path")
    ap.add_argument("--nogui", action="store_true", help="skip GUI even if no file")
    args = ap.parse_args()

    if not args.file and not args.nogui:
        show_gui()
        sys.exit(0)

    try:
        src = open(args.file, encoding="utf-8").read()
    except Exception as e:
        print(f"❌ Error reading {args.file}: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        rows = parse_file(src)
    except Exception as e:
        print(f"❌ Error parsing {args.file}: {e}", file=sys.stderr)
        sys.exit(1)

    print(json.dumps(rows, indent=2))

    swcName = Path(args.file).stem
    write_docx(args.file, swcName, rows, [  # default to all fields
        "Name","Description","Syntax","Triggers","In-Parameters","Out-Parameters",
        "Return Value","Function Type","Inputs","Outputs",
        "Invoked Operations","Used Data Types","Sync/Async","Reentrancy"
    ])

if __name__ == "__main__":
    main()
