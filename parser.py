# #!/usr/bin/env python3
# import re, json, sys, argparse
# from pathlib import Path
# from docx import Document
# from docx.oxml import OxmlElement
# from docx.oxml.ns import qn
# import tkinter as tk
# from tkinter import ttk, filedialog
# import os, platform
# from openpyxl import Workbook, load_workbook
# from openpyxl.styles import Font, PatternFill

# def open_file(path: str):
#     """Open a file with the default OS application."""
#     system = platform.system()
#     if system == "Windows":
#         os.startfile(path)

# def write_excel(file_path: str, rows: list[dict]):
#     """
#     Write out an .xlsx with one sheet named
#     "Runnables and static functions".
#     Headers are styled bold+red font on yellow fill.
#     """
#     headers = [
#       'Name', 'Syntax', 'Return Value', 'In-Parameters', 'Out-Parameters',
#       'Function Type', 'Description', 'Sync/Async', 'Reentrancy',
#       'Triggers', 'Inputs', 'Outputs',
#       'Invoked Operations', 'Used Data Types'
#     ]

#     path = Path(file_path)
#     if path.exists():
#         wb = load_workbook(str(path))
#     else:
#         wb = Workbook()

#     sheet_name = "Runnables and static functions"
#     for name in list(wb.sheetnames):
#         if name != sheet_name:
#             wb.remove(wb[name])
#     if sheet_name in wb.sheetnames:
#         ws = wb[sheet_name]
#     else:
#         ws = wb.create_sheet(title=sheet_name)

#     ws.append(headers)
#     yellow = PatternFill(fill_type="solid", fgColor="FFFFFF00")
#     red    = Font(bold=True, color="FFFF0000")
#     for cell in ws[1]:
#         cell.fill = yellow
#         cell.font = red

#     # append each function’s data
#     for r in rows:
#         ws.append([
#             r.get("name", ""),
#             r.get("syntax", ""),
#             r.get("ret", ""),
#             ", ".join(r.get("inParams", [])),
#             ", ".join(r.get("outParams", [])),
#             r.get("fnType", ""),
#             r.get("description", ""),
#             r.get("Sync_Async", ""),
#             r.get("Reentrancy", ""),
#             r.get("trigger", ""),
#             ", ".join(r.get("inputs", [])),
#             ", ".join(r.get("outputs", [])),
#             ", ".join(r.get("invoked", [])),
#             ", ".join(r.get("used", [])),
#         ])

#     # auto-adjust column widths
#     for col in ws.columns:
#         max_length = 0
#         col_letter = col[0].column_letter
#         for cell in col:
#             try:
#                 val = str(cell.value)
#             except:
#                 val = ""
#             max_length = max(max_length, len(val))
#         ws.column_dimensions[col_letter].width = max_length + 2

#     wb.save(str(path))

# def classify_params(body, params):
#     result = {}
#     for p in params:
#         esc = re.escape(p)
#         ptr_write   = re.search(rf"[\*\(]\s*{esc}\s*\)?\s*(?:[+\-*/]?=|[+\-]{{2}})", body)
#         arrow_write = re.search(rf"\b{esc}\s*->\s*\w+\s*=", body)
#         inc_dec     = re.search(rf"(?:\+\+|--)\s*{esc}\b|\b{esc}\s*(?:\+\+|--)", body)
#         write_api   = re.search(rf"\b\w*(?:Write|Set)\w*\s*\([^;]*\b{esc}\b", body)

#         is_written = bool(ptr_write or arrow_write or inc_dec or write_api)
#         is_read    = bool(re.search(rf"\b{esc}\b", body))

#         result[p] = "INOUT" if (is_written and is_read) else ("OUT" if is_written else "IN")
#     return result

# def get_trigger_comment(comments, pos):
#     """Return last comment block before pos containing a '- triggered on|by …' line."""
#     trig_rx = re.compile(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", re.IGNORECASE)
#     best_end, best_txt = -1, ""
#     for end, txt in comments:
#         if end <= pos and end > best_end and trig_rx.search(txt):
#             best_end, best_txt = end, txt
#     return best_txt

# def show_gui():
#     fields = [
#       "Name","Syntax","Return_Value","In-Parameters","Out-Parameters",
#       "Function_Type","Sync/Async","Reentrancy","Triggers","Inputs",
#       "Outputs","Invoked Operations","Used Data Types", "Description", "Sync/Async", "Reentrancy"
#     ]
#     formats = ["Excel","Word","MD"]

#     root = tk.Tk()
#     root.title("Documentation Slayer")
#     root.configure(bg="#ececec")

#     root.geometry("1000x600")
#     root.minsize(700, 450)

#     # checkbox vars
#     field_vars  = {f: tk.BooleanVar(value=True) for f in fields}
#     format_vars = {f: tk.BooleanVar(value=True) for f in formats}
#     save_dir    = tk.StringVar(value=str(Path.cwd()))

#     frm = ttk.Frame(root, padding=10)
#     frm.pack(fill="both", expand=True)

#     # Fields
#     ttk.Label(frm, text="Select fields:").grid(row=0, column=0, sticky="w")
#     for i, f in enumerate(fields, 1):
#         chk = ttk.Checkbutton(frm, text=f, variable=field_vars[f])
#         chk.grid(row=(i-1)//3+1, column=(i-1)%3, sticky="w", padx=15, pady=5)

#     # Formats
#     ttk.Label(frm, text="Select formats:").grid(row=6, column=0, sticky="w", pady=(10,0))
#     for i, fmt in enumerate(formats, 7):
#         chk = ttk.Checkbutton(frm, text=fmt, variable=format_vars[fmt])
#         chk.grid(row=6 + (i-7)//3+1, column=(i-7)%3, sticky="w", padx=15, pady=5)

#     # Save To…
#     def choose_dir():
#         d = filedialog.askdirectory()
#         if d: save_dir.set(d)
    
#     ttk.Button(frm, text="Save to…", command=choose_dir).grid(row=10, column=0, sticky="w", pady=(10,0))
#     ttk.Label(frm, textvariable=save_dir).grid(row=10, column=1, columnspan=2, sticky="w")

#     def on_run():
#         # collect user choices
#         sel_fields  = [f for f,v in field_vars.items() if v.get()]
#         sel_formats = [f for f,v in format_vars.items() if v.get()]
#         outdir      = Path(save_dir.get())
#         # pick C file
#         cfile = filedialog.askopenfilename(filetypes=[("C files","*.c"),("All","*.*")])
#         if not cfile: return
#         src = open(cfile, encoding="utf-8").read()
#         rows = parse_file(src)
#         stem = Path(cfile).stem
#         # Excel export
#         if "Excel" in sel_formats:
#             xlsx = outdir / f"{stem}.xlsx"
#             write_excel(str(xlsx), rows)
#             open_file(str(xlsx))

#         # Markdown export
#         if "MD" in sel_formats:
#             md = outdir / f"{stem}.md"
#             write_markdown(str(md), rows)
#             open_file(str(md))

#         # Word export
#         if "Word" in sel_formats:
#             docx = outdir / f"{stem}.docx"
#             write_docx(cfile, stem, rows)
#             open_file(str(docx))
#         root.destroy()

#     ttk.Button(frm, text="Run", command=on_run).grid(row=11, column=0, pady=15)
#     root.mainloop()

# def parse_file(src):
#     rows = []
#     reserved = {"if","for","while","switch","do","else","case","sizeof","abs","return"}
#     exclude_invoked = {
#         "VStdLib_MemCpy","VStdLib_MemSet","VStdLib_MemCmp",
#         "memcmp","memcpy","memset","sizeof","abs","return"
#     }

#     # 1) collect comment blocks
#     comments = [(m.end(), m.group(0)) for m in re.finditer(r"/\*[\s\S]*?\*/", src)]

#     # 2a) AUTOSAR runnables: FUNC(return,code) name(
#     runnable_rx = re.compile(
#         r'^[ \t]*FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)\s*'
#         r'([A-Za-z_]\w*)\s*\(',
#         re.MULTILINE
#     )

#     # 2b) static helpers
#     static_rx = re.compile(r'''
#         ^[ \t]*static\s+(?:inline\s+)?                  # static or static inline
#         (?:
#           FUNC\s*\(\s*([A-Za-z_]\w*)\s*,\s*[A-Za-z_]\w*\s*\)  # FUNC macro → group1
#         |
#           ([A-Za-z_]\w*)                                # direct returnType → group2
#         )
#         \s+([A-Za-z_]\w*)\s*\(                          # name → group3
#         ''', re.MULTILINE | re.VERBOSE)

#     # 2c) global helpers
#     global_rx = re.compile(r'''
#         ^[ \t]*(?!static\b)(?!FUNC\b)                   # no static, no FUNC
#         ([A-Za-z_]\w*(?:\s*\*+)?)\s+                    # returnType → group1
#         (?!(?:if|for|while|switch|do|else|case)\b)      # skip reserved
#         ([A-Za-z_]\w*)\s*\(                             # name → group2
#         ''', re.MULTILINE | re.VERBOSE)

#     # signature regex for Syntax field
#     sig_rx = re.compile(r'''
#         ^[ \t]*                                        # indent
#         (?:static\s+)?(?:inline\s+)?                   # optional
#         (?:FUNC\([^)]*\)\s*)?                          # optional FUNC()
#         (?P<ret>[\w\*\s]+?)\s+                         # return type
#         (?P<name>[A-Za-z_]\w*)\s*                      # name
#         \((?P<params>[^)]*)\)                          # params
#         ''', re.MULTILINE | re.VERBOSE)

#     def extract(m, fnType):
#         # pick return type & name
#         if fnType == "Static":
#             retType = m.group(1) or m.group(2)
#             name    = m.group(3)
#         else:
#             retType, name = m.group(1), m.group(2)

#         # find raw_params by counting parentheses
#         L = len(src)
#         depth, i = 1, m.end()
#         while i < L and depth:
#             depth += src[i] == "("
#             depth -= src[i] == ")"
#             i += 1
#         raw_params = src[m.end(): i-1].strip()

#         # build syntax
#         snippet = src[m.start():]
#         sig_m = sig_rx.match(snippet)
#         if sig_m:
#             syntax = f"{sig_m.group('ret').strip()} {sig_m.group('name')}({sig_m.group('params').strip()})"
#         else:
#             syntax = f"{retType} {name}({raw_params})"

#         # skip if no body
#         pos = i
#         while pos < L and (src[pos].isspace() or src.startswith("/*", pos)):
#             pos = src.find("*/", pos)+2 if src.startswith("/*", pos) else pos+1
#         if pos >= L or src[pos] != "{":
#             return

#         # extract body + strip comments
#         brace_idx, depth, j = pos, 1, pos+1
#         while j < L and depth:
#             depth += src[j] == "{"
#             depth -= src[j] == "}"
#             j+=1
#         body = src[brace_idx+1: j-1]
#         code = re.sub(r'/\*[\s\S]*?\*/|//.*', '', body)

#         # triggers
#         if fnType == "Runnable":
#             cm = get_trigger_comment(comments, m.start())
#             trigs = re.findall(r"-\s*triggered\s+(?:on|by)\s+([^\n\r]+)", cm, re.IGNORECASE)
#             trigger = "; ".join(t.strip() for t in trigs)
#         else:
#             trigger = ""

#         # params → names
#         parts = [p.strip() for p in re.split(r",(?![^(]*\))", raw_params) if p.strip()]
#         names = []
#         for p in parts:
#             mm = re.search(r"\b([A-Za-z_]\w*)\s*$", p)
#             names.append(mm.group(1) if mm else p)

#         # IN/OUT
#         dirs = classify_params(body, names)
#         for orig,nm in zip(parts,names):
#             if orig.startswith(("P2CONST(","P2VAR(")):
#                 dirs[nm] = "OUT"
#         inP  = [p for p in names if dirs[p] in ("IN","INOUT")]
#         outP = [p for p in names if dirs[p] in ("OUT","INOUT")]

#         # RTE APIs
#         inputs  = sorted({x.split("(")[0] for x in re.findall(
#                      r"\bRte_(?:Read|DRead|IRead|Receive|IReadRef|IrvRead|IsUpdated|Mode_)[\w_]*\s*\(",
#                      body)})
#         outputs = sorted({x.split("(")[0] for x in re.findall(
#                      r"\bRte_(?:Write|IrvWrite|IWrite|IWriteRef|Switch)[\w_]*\s*\(",
#                      body)})

#         # invoked
#         calls = {x.split("(")[0] for x in re.findall(r"\bRte_Call_[\w_]+\s*\(", code)}
#         plain = re.findall(r"\b([A-Za-z_]\w*)\s*\(", code)
#         locals_ = {c for c in plain
#                    if c not in reserved
#                    and c not in exclude_invoked
#                    and not c.startswith("Rte_")
#                    and c != name}
#         invoked = sorted(c for c in calls|locals_
#                          if not re.fullmatch(r"[A-Z][A-Z0-9_]*", c))

#         # used types
#         used = sorted({t for t in re.findall(
#                       r"\b([A-Za-z_]\w*)\s+[A-Za-z_]\w*\s*(?:[=;])", body)
#                       if t.lower() not in reserved})
        
#         sync_async = ""
#         reentrancy = ""

#         rows.append({
#             "name":      name,
#             "syntax":    syntax,
#             "ret":       retType,
#             "inParams":  inP,
#             "outParams": outP,
#             "fnType":    fnType,
#             "trigger":   trigger,
#             "inputs":    inputs,
#             "outputs":   outputs,
#             "invoked":   invoked,
#             "used":      used,
#             "Sync_Async": sync_async,
#             "Reentrancy": reentrancy
#         })

#     # run extractors
#     for m in runnable_rx.finditer(src):
#         extract(m, "Runnable")
#     for m in static_rx.finditer(src):
#         extract(m, "Static")
#     for m in global_rx.finditer(src):
#         extract(m, "Global")

#     return rows

# def write_docx(source_path, swcName, rows):
#     """Generate a Word .docx per the AUTOSAR‐API template alongside the Excel."""
#     def shade_cell(cell, rgb="9D9D9D"):
#         tcPr = cell._tc.get_or_add_tcPr()
#         shd = OxmlElement('w:shd')
#         shd.set(qn('w:fill'), rgb)
#         tcPr.append(shd)

#     doc = Document()
#     for r in rows:
#         doc.add_paragraph(f"[{r['name']}]", style='Heading 1')
#         tbl = doc.add_table(rows=0, cols=2)
#         tbl.style = 'Table Grid'

#         def add_row(label, value):
#             c0, c1 = tbl.add_row().cells
#             c0.text = label
#             c1.text = value or ""
#             shade_cell(c0)

#         add_row("Service Name",        r['name'])
#         add_row("Syntax",              r['syntax'])
#         add_row("Sync/Async",          "")
#         add_row("Reentrancy",          "")
#         add_row("Parameters (in)",     ", ".join(r['inParams']))
#         add_row("Parameters (inout)",  "")
#         add_row("Parameters (out)",    ", ".join(r['outParams']))
#         add_row("Return value",        r['ret'])
#         add_row("Description",         "")
#         add_row("Triggers",            r['trigger'])
#         add_row("Inputs",              ", ".join(r['inputs']))
#         add_row("Outputs",             ", ".join(r['outputs']))
#         add_row("Invoked Operations",  ", ".join(r['invoked']))
#         add_row("Used Data Types",     ", ".join(r['used']))
#         add_row("Available via",       "")
#         doc.add_page_break()

#     # save alongside the Excel (.xlsx) in the same directory
#     excel_path = Path(source_path).with_suffix('.xlsx')
#     docx_path  = excel_path.with_suffix('.docx')
#     doc.save(docx_path)

#     # print is shitty here don't use it -> will ruin the parsing of the output JSON file.
#     # print(f"↪️  Written Word doc: {docx_path}", file=sys.stderr)

# def write_markdown(file_path: str, rows: list[dict]):
#     """
#     rows: list of dicts with keys:
#       name, syntax, Sync_Async, Reentrancy, ret, inParams, outParams,
#       fnType, trigger, inputs, outputs, invoked, used, (optional) description
#     """
#     lines = []
#     for r in rows:
#         lines.append(f"## {r['name']}\n")
#         lines.append("| Field | Value |")
#         lines.append("|-------|-------|")
#         lines.append(f"| Syntax | `{r['syntax']}` |")
#         lines.append(f"| Sync/Async | `{r['Sync_Async']}` |")
#         lines.append(f"| Reentrancy | `{r['Reentrancy']}` |")
#         lines.append(f"| Return Value | `{r['ret']}` |")
#         lines.append(f"| In-Parameters | {', '.join(r['inParams'])} |")
#         lines.append(f"| Out-Parameters | {', '.join(r['outParams'])} |")
#         lines.append(f"| Function Type | {r['fnType']} |")
#         lines.append(f"| Triggers | {r['trigger']} |")
#         lines.append(f"| Inputs | {', '.join(r['inputs'])} |")
#         lines.append(f"| Outputs | {', '.join(r['outputs'])} |")
#         lines.append(f"| Invoked Operations | {', '.join(r['invoked'])} |")
#         lines.append(f"| Used Data Types | {', '.join(r['used'])} |")
#         # use empty string if no description key
#         desc = r.get('description', '')
#         lines.append(f"| Description | {desc} |")
#         lines.append("")  # blank line between entries

#     # write out the file
#     with open(file_path, 'w', encoding='utf-8') as f:
#         f.write("\n".join(lines))


# def main():
#     ap = argparse.ArgumentParser(description="Extract and document Runnables")
#     # make the source‐file argument optional
#     ap.add_argument("file", nargs="?", default=None, help="C source file path")
#     ap.add_argument("--nogui", action="store_true", help="skip GUI even if no file")
#     args = ap.parse_args()

#     # no file → show GUI
#     if not args.file and not args.nogui:
#         # no args → interactive GUI
#         show_gui()
#         sys.exit(0)

#     try:
#         src = open(args.file, encoding="utf-8").read()
#     except Exception as e:
#         print(f"❌ Error reading {args.file}: {e}", file=sys.stderr)
#         sys.exit(1)

#     try:
#         rows = parse_file(src)
#     except Exception as e:
#         print(f"❌ Error parsing {args.file}: {e}", file=sys.stderr)
#         sys.exit(1)

#     # output JSON for compatibility
#     print(json.dumps(rows, indent=2))

#     # write the Word doc next to the Excel
#     swcName = Path(args.file).stem
#     write_docx(args.file, swcName, rows)

# if __name__ == "__main__":
#     main()

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
      'Syntax':        lambda r: f"`{r['syntax']}`",
      'Sync/Async':    lambda r: f"`{r['Sync_Async']}`",
      'Reentrancy':    lambda r: f"`{r['Reentrancy']}`",
      'Return Value':  lambda r: f"`{r['ret']}`",
      'In-Parameters': lambda r: ", ".join(r['inParams']),
      'Out-Parameters':lambda r: ", ".join(r['outParams']),
      'Function Type': lambda r: r['fnType'],
      'Description':   lambda r: r.get('description', ''),
      'Triggers':      lambda r: r['trigger'],
      'Inputs':        lambda r: ", ".join(r['inputs']),
      'Outputs':       lambda r: ", ".join(r['outputs']),
      'Invoked Operations': lambda r: ", ".join(r['invoked']),
      'Used Data Types':    lambda r: ", ".join(r['used']),
    }

    lines = []
    for r in rows:
        lines.append(f"## {r['name']}\n")
        lines.append("| Field | Value |")
        lines.append("|-------|-------|")
        for label in sel_fields:
            if label == 'Name':
                continue  # Name is the header, not a row field
            getter = FIELD_GETTERS.get(label)
            if getter:
                lines.append(f"| {label} | {getter(r)} |")
        lines.append("")  # blank line between functions

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))

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
    root.geometry("900x550")
    root.minsize(700, 450)

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
        sel_fields  = [f for f,v in field_vars.items() if v.get()]
        sel_formats = [f for f,v in format_vars.items() if v.get()]
        outdir      = Path(save_dir.get())
        cfile = filedialog.askopenfilename(filetypes=[("C files","*.c"),("All","*.*")])
        if not cfile:
            return
        src   = open(cfile, encoding="utf-8").read()
        rows  = parse_file(src)
        stem  = Path(cfile).stem

        if "Excel" in sel_formats:
            xlsx = outdir / f"{stem}.xlsx"
            write_excel(str(xlsx), rows, sel_fields)
            open_file(str(xlsx))

        if "MD" in sel_formats:
            md = outdir / f"{stem}.md"
            write_markdown(str(md), rows, sel_fields)
            open_file(str(md))

        if "Word" in sel_formats:
            docx = outdir / f"{stem}.docx"
            write_docx(cfile, stem, rows, sel_fields)
            # open_file(str(docx))

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
