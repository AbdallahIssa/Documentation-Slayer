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
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import subprocess


"""
I'm using This command to get the executable for the parser script:
pyinstaller --onefile --windowed --name "Doc-Slayer" --icon "vehiclevo_logo_Basic.ico" --add-data "vehiclevo_logo_Basic.ico;." --add-data "CodeSmasher.exe;." parser.py
"""

class PasswordManager:
    """Manages password authentication with session persistence"""
    def __init__(self):
        self.is_authenticated = False
        self.attempts = 0
        self.max_attempts = 3
    
    def reset(self):
        """Reset authentication state"""
        self.is_authenticated = False
        self.attempts = 0


# Global password manager instance
password_manager = PasswordManager()


def ask_password(parent_window=None):
    """Ask for password to access Activity Diagram tab with 3 attempts"""
    if password_manager.is_authenticated:
        return True
    
    if password_manager.attempts >= password_manager.max_attempts:
        messagebox.showerror("Access Denied", "Maximum password attempts exceeded. Application will close.")
        if parent_window:
            parent_window.destroy()
        sys.exit(1)
    
    password_window = tk.Toplevel(parent_window)
    password_window.title("Access Activity Diagram")
    password_window.geometry("350x200")
    password_window.resizable(False, False)
    password_window.grab_set()  # Make it modal
    
    # Center the window
    password_window.update_idletasks()
    w = password_window.winfo_width()
    h = password_window.winfo_height()
    sw = password_window.winfo_screenwidth()
    sh = password_window.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    password_window.geometry(f"{w}x{h}+{x}+{y}")
    
    # Add icon to password window
    if getattr(sys, "frozen", False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).parent
    icon_path = base_path / "vehiclevo_logo_Basic.ico"
    try:
        password_window.iconbitmap(str(icon_path))
    except Exception:
        pass
    
    ttk.Label(password_window, text="Enter password to access Activity Diagram:").pack(pady=10)
    
    attempts_left = password_manager.max_attempts - password_manager.attempts
    ttk.Label(password_window, text=f"Attempts remaining: {attempts_left}").pack(pady=5)
    
    password_var = tk.StringVar()
    password_entry = ttk.Entry(password_window, textvariable=password_var, show="*", width=30)
    password_entry.pack(pady=5)
    password_entry.focus()
    
    result = {"access_granted": False}
    
    def check_password():
        password_manager.attempts += 1
        
        if password_var.get() == "Vehiclevo@1234":
            password_manager.is_authenticated = True
            result["access_granted"] = True
            password_window.destroy()
        else:
            if password_manager.attempts >= password_manager.max_attempts:
                messagebox.showerror("Access Denied", "Maximum password attempts exceeded. Application will close.")
                password_window.destroy()
                if parent_window:
                    parent_window.destroy()
                sys.exit(1)
            else:
                remaining = password_manager.max_attempts - password_manager.attempts
                error_msg = f"Wrong password! {remaining} attempt(s) remaining."
                messagebox.showerror("Invalid Password", error_msg)
                password_entry.delete(0, tk.END)
                password_window.destroy()
    
    def on_enter(event):
        check_password()
    
    def on_cancel():
        password_window.destroy()
    
    password_entry.bind('<Return>', on_enter)
    
    button_frame = ttk.Frame(password_window)
    button_frame.pack(pady=10)
    
    ttk.Button(button_frame, text="OK", command=check_password).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=5)
    
    password_window.protocol("WM_DELETE_WINDOW", on_cancel)
    password_window.wait_window()
    
    return result["access_granted"]


def open_file(path: str):
    """Open a file with the default OS application."""
    system = platform.system()
    if system == "Windows":
        os.startfile(path)
    elif system == "Darwin":
        os.system(f'open "{path}"')
    else:
        os.system(f'xdg-open "{path}"')

def write_excel(file_path: str, functions: list[dict], macros: list[dict], variables: list[dict], 
                sel_function_fields: list[str], sel_macro_fields: list[str], sel_variable_fields: list[str]):
    """
    Write an .xlsx with three sheets: Functions, Macros, and Variables.
    Only the selected fields columns are written for each sheet.
    Headers are styled bold+red font on yellow fill.
    """
    FUNCTION_HEADERS = [
      'Line Number', 'Name', 'Syntax', 'Return Value', 'In-Parameters', 'Out-Parameters',
      'Function Type', 'Description', 'Sync/Async', 'Reentrancy',
      'Triggers', 'Inputs', 'Outputs',
      'Invoked Operations', 'Used Data Types'
    ]

    MACRO_HEADERS = [
        'Line Number', 'Name', 'Value'
    ]

    VARIABLE_HEADERS = [
        'Line Number', 'Name', 'Data Type', 'Initial Value', 'Scope'
    ]

    path = Path(file_path)
    if path.exists():
        wb = load_workbook(str(path))
        # Remove all existing sheets
        for name in list(wb.sheetnames):
            wb.remove(wb[name])
    else:
        wb = Workbook()
        # Remove default sheet
        wb.remove(wb.active)

    yellow = PatternFill(fill_type="solid", fgColor="FFFFFF00")
    red = Font(bold=True, color="FFFF0000")

    # Functions sheet
    if functions:
        ws_functions = wb.create_sheet(title="Runnables and static functions")
        function_headers = [h for h in FUNCTION_HEADERS if h in sel_function_fields]
        ws_functions.append(function_headers)
        
        for cell in ws_functions[1]:
            cell.fill = yellow
            cell.font = red

        for r in functions:
            row = []
            for h in function_headers:
                if h == 'Line Number':
                    row.append(r.get('lineNumber', ''))
                elif h == 'Name':
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
            ws_functions.append(row)

        # Auto-adjust column widths
        for col in ws_functions.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = str(cell.value or "")
                max_length = max(max_length, len(val))
            ws_functions.column_dimensions[col_letter].width = max_length + 2

    # Macros sheet
    if macros:
        ws_macros = wb.create_sheet(title="Macros")
        macro_headers = [h for h in MACRO_HEADERS if h in sel_macro_fields]
        ws_macros.append(macro_headers)
        
        for cell in ws_macros[1]:
            cell.fill = yellow
            cell.font = red

        for r in macros:
            row = []
            for h in macro_headers:
                if h == 'Line Number':
                    row.append(r.get('lineNumber', ''))
                elif h == 'Name':
                    row.append(r['name'])
                elif h == 'Value':
                    row.append(r['value'])
            ws_macros.append(row)

        # Auto-adjust column widths
        for col in ws_macros.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = str(cell.value or "")
                max_length = max(max_length, len(val))
            ws_macros.column_dimensions[col_letter].width = max_length + 2

    # Variables sheet
    if variables:
        ws_variables = wb.create_sheet(title="Variables")
        variable_headers = [h for h in VARIABLE_HEADERS if h in sel_variable_fields]
        ws_variables.append(variable_headers)
        
        for cell in ws_variables[1]:
            cell.fill = yellow
            cell.font = red

        for r in variables:
            row = []
            for h in variable_headers:
                if h == 'Line Number':
                    row.append(r.get('lineNumber', ''))
                elif h == 'Name':
                    row.append(r['name'])
                elif h == 'Data Type':
                    row.append(r['dataType'])
                elif h == 'Initial Value':
                    row.append(r['initialValue'])
                elif h == 'Scope':
                    row.append(r['scope'])
            ws_variables.append(row)

        # Auto-adjust column widths
        for col in ws_variables.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = str(cell.value or "")
                max_length = max(max_length, len(val))
            ws_variables.column_dimensions[col_letter].width = max_length + 2

    wb.save(str(path))

def write_markdown(file_path: str, functions: list[dict], macros: list[dict], variables: list[dict],
                   sel_function_fields: list[str], sel_macro_fields: list[str], sel_variable_fields: list[str]):
    """
    Write a Markdown file with tables for functions, macros, and variables.
    Only the selected fields are included in each table.
    """
    FUNCTION_FIELD_GETTERS = {
      'Line Number':      lambda r: str(r.get('lineNumber', '')),
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

    MACRO_FIELD_GETTERS = {
        'Line Number': lambda r: str(r.get('lineNumber', '')),
        'Name':  lambda r: r['name'],
        'Value': lambda r: f"`{r['value']}`",
    }

    VARIABLE_FIELD_GETTERS = {
        'Line Number':   lambda r: str(r.get('lineNumber', '')),
        'Name':         lambda r: r['name'],
        'Data Type':    lambda r: f"`{r['dataType']}`",
        'Initial Value': lambda r: f"`{r['initialValue']}`",
        'Scope':        lambda r: r['scope'],
    }

    lines = []

    # Functions section
    if functions:
        lines.append("# Functions")
        lines.append(f"**Selected fields:** {', '.join(sel_function_fields)}")
        lines.append("")

        for r in functions:
            lines.append(f"## {r['name']}")
            lines.append("")
            lines.append("| Field | Value |")
            lines.append("|-------|-------|")
            for label in sel_function_fields:
                getter = FUNCTION_FIELD_GETTERS.get(label)
                if getter:
                    value = getter(r)
                    lines.append(f"| {label} | {value} |")
            lines.append("")

    # Macros section
    if macros:
        lines.append("# Macros")
        lines.append(f"**Selected fields:** {', '.join(sel_macro_fields)}")
        lines.append("")

        for r in macros:
            lines.append(f"## {r['name']}")
            lines.append("")
            lines.append("| Field | Value |")
            lines.append("|-------|-------|")
            for label in sel_macro_fields:
                getter = MACRO_FIELD_GETTERS.get(label)
                if getter:
                    value = getter(r)
                    lines.append(f"| {label} | {value} |")
            lines.append("")

    # Variables section
    if variables:
        lines.append("# Variables")
        lines.append(f"**Selected fields:** {', '.join(sel_variable_fields)}")
        lines.append("")

        for r in variables:
            lines.append(f"## {r['name']}")
            lines.append("")
            lines.append("| Field | Value |")
            lines.append("|-------|-------|")
            for label in sel_variable_fields:
                getter = VARIABLE_FIELD_GETTERS.get(label)
                if getter:
                    value = getter(r)
                    lines.append(f"| {label} | {value} |")
            lines.append("")

    Path(file_path).write_text("\n".join(lines), encoding="utf-8")

def write_docx(source_path: str, swcName: str, functions: list[dict], macros: list[dict], variables: list[dict],
               sel_function_fields: list[str], sel_macro_fields: list[str], sel_variable_fields: list[str]):
    """
    Generate a Word .docx per AUTOSAR template.
    Includes functions, macros, and variables with only selected fields.
    """
    def shade_cell(cell, rgb="9D9D9D"):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), rgb)
        tcPr.append(shd)

    doc = Document()

    # Functions section
    if functions:
        doc.add_heading('Functions', level=1)
        
        for r in functions:
            doc.add_paragraph(f"[{r['name']}]", style='Heading 2')
            tbl = doc.add_table(rows=0, cols=2)
            tbl.style = 'Table Grid'

            def add_row(label, value):
                c0, c1 = tbl.add_row().cells
                c0.text = label
                c1.text = value or ""
                shade_cell(c0)

            if 'Line Number' in sel_function_fields:
                add_row("Line Number", str(r.get('lineNumber', '')))
            if 'Name' in sel_function_fields:
                add_row("Service Name", r['name'])
            if 'Syntax' in sel_function_fields:
                add_row("Syntax", r['syntax'])
            if 'Sync/Async' in sel_function_fields:
                add_row("Sync/Async", r['Sync_Async'])
            if 'Reentrancy' in sel_function_fields:
                add_row("Reentrancy", r['Reentrancy'])
            if 'In-Parameters' in sel_function_fields:
                add_row("Parameters (in)", ", ".join(r['inParams']))
            if 'Out-Parameters' in sel_function_fields:
                add_row("Parameters (out)", ", ".join(r['outParams']))
            if 'Function Type' in sel_function_fields:
                add_row("Function Type", r['fnType'])
            if 'Description' in sel_function_fields:
                add_row("Description", r.get('description', ''))
            if 'Triggers' in sel_function_fields:
                add_row("Triggers", r['trigger'])
            if 'Inputs' in sel_function_fields:
                add_row("Inputs", ", ".join(r['inputs']))
            if 'Outputs' in sel_function_fields:
                add_row("Outputs", ", ".join(r['outputs']))
            if 'Invoked Operations' in sel_function_fields:
                add_row("Invoked Operations", ", ".join(r['invoked']))
            if 'Used Data Types' in sel_function_fields:
                add_row("Used Data Types", ", ".join(r['used']))

            doc.add_page_break()

    # Macros section
    if macros:
        doc.add_heading('Macros', level=1)
        
        for r in macros:
            doc.add_paragraph(f"[{r['name']}]", style='Heading 2')
            tbl = doc.add_table(rows=0, cols=2)
            tbl.style = 'Table Grid'

            def add_row(label, value):
                c0, c1 = tbl.add_row().cells
                c0.text = label
                c1.text = value or ""
                shade_cell(c0)

            if 'Line Number' in sel_macro_fields:
                add_row("Line Number", str(r.get('lineNumber', '')))
            if 'Name' in sel_macro_fields:
                add_row("Macro Name", r['name'])
            if 'Value' in sel_macro_fields:
                add_row("Value", r['value'])

            doc.add_page_break()

    # Variables section
    if variables:
        doc.add_heading('Variables', level=1)
        
        for r in variables:
            doc.add_paragraph(f"[{r['name']}]", style='Heading 2')
            tbl = doc.add_table(rows=0, cols=2)
            tbl.style = 'Table Grid'

            def add_row(label, value):
                c0, c1 = tbl.add_row().cells
                c0.text = label
                c1.text = value or ""
                shade_cell(c0)

            if 'Line Number' in sel_variable_fields:
                add_row("Line Number", str(r.get('lineNumber', '')))
            if 'Name' in sel_variable_fields:
                add_row("Variable Name", r['name'])
            if 'Data Type' in sel_variable_fields:
                add_row("Data Type", r['dataType'])
            if 'Initial Value' in sel_variable_fields:
                add_row("Initial Value", r['initialValue'])
            if 'Scope' in sel_variable_fields:
                add_row("Scope", r['scope'])

            doc.add_page_break()

    excel_path = Path(source_path).with_suffix('.xlsx')
    docx_path = excel_path.with_suffix('.docx')
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

def parse_macros(src: str) -> list[dict]:
    """Extract #define macros from C source code."""
    # Regex for a macro header: #define NAME [optional(param,list)] body-fragment
    HEADER_RE = re.compile(r"^\s*#define\s+(\w+)\s*(\([^)]*\))?\s*(.*)", re.MULTILINE)

    def should_skip(body: str) -> bool:
        """Return True if the macro should be excluded."""
        stripped = body.strip()
        if not stripped:
            return True
        return False

    lines = src.split('\n')
    macros = []

    i = 0
    while i < len(lines):
        m = HEADER_RE.match(lines[i])
        if not m:
            i += 1
            continue

        line_number = i + 1  # Track line number (1-indexed)
        name = m.group(1)              # macro name
        paramlist = m.group(2)         # None if object-like macro
        first = m.group(3).rstrip()    # first chunk of body

        # Gather continuation lines ending in "\"
        body_parts = [first]
        while body_parts and body_parts[-1].endswith("\\"):
            body_parts[-1] = body_parts[-1][:-1].rstrip()   # drop trailing "\"
            i += 1
            if i >= len(lines):
                break
            body_parts.append(lines[i].rstrip())

        body_text = " ".join(part.strip() for part in body_parts).strip()

        if not should_skip(body_text):
            macros.append({
                "name": name,
                "value": body_text,
                "lineNumber": line_number
            })

        i += 1

    return macros

def parse_variables(src: str) -> list[dict]:
    """Extract global and static global variables from C source code."""
    variables = []
    
    # Remove preprocessor directives and comments to avoid false matches
    src_clean = re.sub(r'^\s*#.*$', '', src, flags=re.MULTILINE)
    src_clean = re.sub(r'/\*[\s\S]*?\*/|//.*', '', src_clean)
    
    # First, identify all function bodies to exclude local variables
    function_bodies = []
    
    # Find all functions (including FUNC macros, static, and global functions)
    func_patterns = [
        re.compile(r'FUNC\s*\([^)]*\)\s*[A-Za-z_]\w*\s*\([^)]*\)\s*\{', re.MULTILINE),
        re.compile(r'(?:static\s+)?(?:inline\s+)?[A-Za-z_]\w*(?:\s*\*+)?\s+[A-Za-z_]\w*\s*\([^)]*\)\s*\{', re.MULTILINE)
    ]
    
    for pattern in func_patterns:
        for match in pattern.finditer(src_clean):
            # Find the complete function body
            start_pos = match.end() - 1  # position of opening '{'
            brace_count = 1
            pos = start_pos + 1
            
            while pos < len(src_clean) and brace_count > 0:
                if src_clean[pos] == '{':
                    brace_count += 1
                elif src_clean[pos] == '}':
                    brace_count -= 1
                pos += 1
            
            if brace_count == 0:
                function_bodies.append((start_pos, pos))
            else:
                 # If we reached end of file with unclosed braces, don't include this as a function body
                 # This prevents the entire end of file from being marked as "inside function"
                 pass
    
    def is_inside_function(position):
        """Check if a position is inside any function body."""
        for start, end in function_bodies:
            if start <= position <= end:
                return True
        return False
    
    # Debug: Print function bodies for troubleshooting
    # print(f"Function bodies detected: {function_bodies}")

    # Pattern for extern variable declarations
    extern_var_pattern = re.compile(r'''
        ^[ \t]*                                    # start of line, optional whitespace
        (extern\s+)                                # extern keyword
        ([A-Za-z_]\w*(?:\s*\*+)?)                 # data type (with optional pointers)
        \s+                                        # whitespace
        ([A-Za-z_]\w*)                            # variable name
        \s*;(?=\s|$)                                       # semicolon (extern vars don't have initialization)
        ''', re.MULTILINE | re.VERBOSE)
    
    # Pattern for static/regular variable declarations
    static_var_pattern = re.compile(r'''
        ^[ \t]*                                    # start of line, optional whitespace
        (static\s+)?                               # optional 'static' keyword
        ([A-Za-z_]\w*(?:\s*\*+)?)                 # data type (with optional pointers)
        \s+                                        # whitespace
        ([A-Za-z_]\w*)                            # variable name
        (?:\s*=\s*([^;]+))?                       # optional initialization
        \s*;(?=\s|$)                                       # semicolon
        ''', re.MULTILINE | re.VERBOSE)
    
    # Array pattern: type name[size] = {...};
    static_array_pattern = re.compile(r'''
        ^[ \t]*                                    # start of line, optional whitespace
        (static\s+)?                               # optional 'static' keyword
        ([A-Za-z_]\w*(?:\s*\*+)?)                 # data type
        \s+                                        # whitespace
        ([A-Za-z_]\w*)                            # variable name
        \s*\[([^\]]*)\]                           # array brackets with size
        (?:\s*=\s*\{([^}]*)\})?                   # optional array initialization
        \s*;(?=\s|$)                                       # semicolon
        ''', re.MULTILINE | re.VERBOSE)
    
    # Keywords to exclude (function keywords, control flow, etc.)
    exclude_keywords = {
        'if', 'for', 'while', 'switch', 'do', 'else', 'case', 'return',
        'goto', 'break', 'continue', 'sizeof', 'typedef', 'struct', 'union',
        'enum', 'extern', 'register', 'auto', 'volatile', 'const', 'inline',
        'unsigned', 'signed', 'long', 'short', 'void', 'char', 'int', 'float', 'double'
    }
    
    # Find extern variable declarations
    for match in extern_var_pattern.finditer(src_clean):
        if is_inside_function(match.start()):
            continue

        extern_kw = match.group(1)
        data_type = match.group(2).strip()
        var_name = match.group(3)

        if extern_kw and var_name not in exclude_keywords and data_type.lower() not in exclude_keywords:
            line_number = get_line_number(src_clean, match.start())
            variables.append({
                "name": var_name,
                "dataType": data_type,
                "initialValue": "",
                "scope": "Extern",
                "lineNumber": line_number
            })
    
    # Find static/regular variable declarations
    for match in static_var_pattern.finditer(src_clean):
        if is_inside_function(match.start()):
            continue

        static_kw = match.group(1)
        data_type = match.group(2).strip()
        var_name = match.group(3)
        init_value = match.group(4).strip() if match.group(4) else ""

        if var_name in exclude_keywords or data_type.lower() in exclude_keywords:
            continue

        # Check if this looks like a function (has parentheses after the name)
        next_pos = match.end()
        if next_pos < len(src_clean):
            remaining = src_clean[next_pos:next_pos+50]
            if '(' in remaining.split(';')[0]:
                continue

        # Skip if it looks like a function pointer or typedef
        full_match = match.group(0)
        if '(*' in full_match or 'typedef' in full_match:
            continue

        scope = "Static Global" if static_kw else "Global"
        line_number = get_line_number(src_clean, match.start())

        variables.append({
            "name": var_name,
            "dataType": data_type,
            "initialValue": init_value,
            "scope": scope,
            "lineNumber": line_number
        })

    # Find array declarations
    for match in static_array_pattern.finditer(src_clean):
        if is_inside_function(match.start()):
            continue

        static_kw = match.group(1)
        data_type = match.group(2).strip()
        var_name = match.group(3)
        array_size = match.group(4).strip() if match.group(4) else ""
        init_value = match.group(5).strip() if match.group(5) else ""

        if var_name in exclude_keywords or data_type.lower() in exclude_keywords:
            continue

        scope = "Static Global" if static_kw else "Global"
        line_number = get_line_number(src_clean, match.start())

        full_type = f"{data_type}[{array_size}]"

        variables.append({
            "name": var_name,
            "dataType": full_type,
            "initialValue": init_value,
            "scope": scope,
            "lineNumber": line_number
        })

    # Post-processing: Clean up any remaining preprocessor keywords from datatypes
    preprocessor_keywords = {
        'endif', 'if', 'ifdef', 'ifndef', 'else', 'elif', 'define', 'include', 
        'undef', 'pragma', 'warning', 'error', 'line'
    }
    
    cleaned_variables = []
    for var in variables:
        # Check if dataType contains any preprocessor keywords
        datatype_words = var['dataType'].lower().split()
        if not any(word in preprocessor_keywords for word in datatype_words):
            cleaned_variables.append(var)
        # Optionally, I can clean individual words instead of removing entire variable:
        # cleaned_datatype = ' '.join(word for word in var['dataType'].split() 
        #                            if word.lower() not in preprocessor_keywords)
        # if cleaned_datatype.strip():
        #     var['dataType'] = cleaned_datatype.strip()
        #     cleaned_variables.append(var)
    
    return cleaned_variables

def get_line_number(src: str, pos: int) -> int:
    """Convert string position to line number (1-indexed)."""
    return src[:pos].count('\n') + 1

def parse_file(src: str) -> tuple[list, list, list]:
    """Parse file and return (functions, macros, variables)."""
    functions = []
    reserved = {"if","for","while","switch","do","else","case","sizeof","abs","return", "endif"}
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

        # Calculate line number
        line_number = get_line_number(src, m.start())

        functions.append({
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
            "description": description,
            "lineNumber": line_number
        })

    for m in runnable_rx.finditer(src):
        extract(m, "Runnable")
    for m in static_rx.finditer(src):
        extract(m, "Static")
    for m in global_rx.finditer(src):
        extract(m, "Global")

    # Parse macros and variables
    macros = parse_macros(src)
    variables = parse_variables(src)

    return functions, macros, variables

def show_gui():
    function_fields = [
      "Line Number", "Name", "Description", "Syntax", "Triggers", "In-Parameters", "Out-Parameters",
      "Return Value", "Function Type", "Inputs", "Outputs",
      "Invoked Operations", "Used Data Types", "Sync/Async", "Reentrancy"
    ]

    macro_fields = [
        "Line Number", "Name", "Value"
    ]

    variable_fields = [
        "Line Number", "Name", "Data Type", "Initial Value", "Scope"
    ]

    formats = ["Excel", "Word", "MD"]

    root = tk.Tk()
    root.title("Documentation Slayer")
    root.configure(bg="#ececec")
    if getattr(sys, "frozen", False):
        # running from PyInstaller bundle
        base_path = Path(sys._MEIPASS)
    else:
        # running as a normal script
        base_path = Path(__file__).parent
    icon_path = base_path / "vehiclevo_logo_Basic.ico"
    try:
        root.iconbitmap(str(icon_path))
    except Exception:
        pass  # skip if missing or invalid
    root.geometry("900x600")
    root.minsize(900, 600)
    root.maxsize(900, 600)
    root.resizable(False, False) # no resizing

    root.update_idletasks()         # ensure winfo_width/height are accurate
    w = root.winfo_width()
    h = root.winfo_height()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

    # Create notebook for tabs
    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    # Functions tab
    functions_frame = ttk.Frame(notebook)
    notebook.add(functions_frame, text="Functions")

    function_vars = {f: tk.BooleanVar(value=(f != "Line Number")) for f in function_fields}

    ttk.Label(functions_frame, text="Select function fields:").grid(row=0, column=0, sticky="w", padx=10, pady=5)

    # Select All / Deselect All buttons
    button_frame_func = ttk.Frame(functions_frame)
    button_frame_func.grid(row=0, column=1, columnspan=2, sticky="e", padx=10, pady=5)

    def select_all_functions():
        for var in function_vars.values():
            var.set(True)

    def deselect_all_functions():
        for var in function_vars.values():
            var.set(False)

    ttk.Button(button_frame_func, text="Select All", command=select_all_functions).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame_func, text="Deselect All", command=deselect_all_functions).pack(side=tk.LEFT, padx=5)

    for i, f in enumerate(function_fields):
        chk = ttk.Checkbutton(functions_frame, text=f, variable=function_vars[f])
        chk.grid(row=(i//3)+1, column=i%3, sticky="w", padx=15, pady=2)

    # Macros tab
    macros_frame = ttk.Frame(notebook)
    notebook.add(macros_frame, text="Macros")

    macro_vars = {f: tk.BooleanVar(value=(f != "Line Number")) for f in macro_fields}

    ttk.Label(macros_frame, text="Select macro fields:").grid(row=0, column=0, sticky="w", padx=10, pady=5)

    # Select All / Deselect All buttons
    button_frame_macro = ttk.Frame(macros_frame)
    button_frame_macro.grid(row=0, column=1, columnspan=2, sticky="e", padx=10, pady=5)

    def select_all_macros():
        for var in macro_vars.values():
            var.set(True)

    def deselect_all_macros():
        for var in macro_vars.values():
            var.set(False)

    ttk.Button(button_frame_macro, text="Select All", command=select_all_macros).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame_macro, text="Deselect All", command=deselect_all_macros).pack(side=tk.LEFT, padx=5)

    for i, f in enumerate(macro_fields):
        chk = ttk.Checkbutton(macros_frame, text=f, variable=macro_vars[f])
        chk.grid(row=(i//3)+1, column=i%3, sticky="w", padx=15, pady=2)

    # Variables tab
    variables_frame = ttk.Frame(notebook)
    notebook.add(variables_frame, text="Variables")

    variable_vars = {f: tk.BooleanVar(value=(f != "Line Number")) for f in variable_fields}

    ttk.Label(variables_frame, text="Select variable fields:").grid(row=0, column=0, sticky="w", padx=10, pady=5)

    # Select All / Deselect All buttons
    button_frame_var = ttk.Frame(variables_frame)
    button_frame_var.grid(row=0, column=1, columnspan=2, sticky="e", padx=10, pady=5)

    def select_all_variables():
        for var in variable_vars.values():
            var.set(True)

    def deselect_all_variables():
        for var in variable_vars.values():
            var.set(False)

    ttk.Button(button_frame_var, text="Select All", command=select_all_variables).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame_var, text="Deselect All", command=deselect_all_variables).pack(side=tk.LEFT, padx=5)

    for i, f in enumerate(variable_fields):
        chk = ttk.Checkbutton(variables_frame, text=f, variable=variable_vars[f])
        chk.grid(row=(i//3)+1, column=i%3, sticky="w", padx=15, pady=2)

    # Activity Diagram tab (password protected)
    activity_frame = ttk.Frame(notebook)
    
    def on_activity_tab_selected(event):
        selected_tab = event.widget.tab('current')['text']
        if selected_tab == "Activity Diagram" and not password_manager.is_authenticated:
            if not ask_password(root):
                # If password wrong or cancelled, go back to Variables tab
                notebook.select(2)  # Variables tab index
                return
    
    notebook.bind("<<NotebookTabChanged>>", on_activity_tab_selected)
    notebook.add(activity_frame, text="Activity Diagram")
    
    # Activity Diagram content with button instead of checkbox
    ttk.Label(activity_frame, text="Activity Diagram Generator", font=("Arial", 14, "bold")).pack(pady=20)
    ttk.Label(activity_frame, text="Click the button below to generate activity diagrams from your C source file.").pack(pady=10)
    
    
    def generate_activity_diagram():
        if not password_manager.is_authenticated:
            if not ask_password(root):
                return
        
        # # Ask for C file if not already selected
        # if not selected_file["path"]:
        #     cfile = filedialog.askopenfilename(
        #         title="Select C source file for activity diagram",
        #         filetypes=[("C files","*.c"), ("All files","*.*")]
        #     )
        #     if not cfile:
        #         return
        #     selected_file["path"] = cfile
        
        # # Get output directory
        # outdir = Path(save_dir.get())
        # stem = Path(selected_file["path"]).stem
        
        # Execute the external CodeSmasher.exe
        try:
            if getattr(sys, "frozen", False):
                base_path = Path(sys._MEIPASS)
            else:
                base_path = Path(__file__).parent
            
            exe_path = base_path / "CodeSmasher.exe"
            if exe_path.exists():
                # Run CodeSmasher.exe 
                cmd = [str(exe_path)]
                result = subprocess.run(cmd, capture_output=True, text=True)
                if result.returncode == 0:
                    messagebox.showinfo("Success", "Activity diagrams generated successfully!")
                else:
                    messagebox.showerror("Error", f"Failed to generate activity diagrams:\n{result.stderr}")
            else:
                messagebox.showerror("Error", "CodeSmasher.exe not found! Please ensure it's in the same directory as this script.")
        except Exception as e:
            messagebox.showerror("Error", f"Error running CodeSmasher.exe: {str(e)}")
    
    generate_btn = ttk.Button(activity_frame, text="Generate Activity Diagram", command=generate_activity_diagram, style="Accent.TButton")
    generate_btn.pack(pady=30)
    
    # Add style for accent button
    style = ttk.Style()
    style.configure("Accent.TButton", font=("Arial", 11, "bold"))

    # Settings frame (at bottom)
    settings_frame = ttk.Frame(root)
    settings_frame.pack(fill="x", padx=10, pady=(0, 10))

    format_vars = {f: tk.BooleanVar(value=True) for f in formats}
    save_dir = tk.StringVar(value=str(Path.cwd()))

    ttk.Label(settings_frame, text="Select formats:").grid(row=0, column=0, sticky="w", pady=5)
    for i, fmt in enumerate(formats):
        chk = ttk.Checkbutton(settings_frame, text=fmt, variable=format_vars[fmt])
        chk.grid(row=0, column=i+1, sticky="w", padx=15, pady=5)

    def choose_dir():
        d = filedialog.askdirectory()
        if d:
            save_dir.set(d)
    ttk.Button(settings_frame, text="Save to…", command=choose_dir).grid(row=1, column=0, sticky="w", pady=5)
    ttk.Label(settings_frame, textvariable=save_dir).grid(row=1, column=1, columnspan=3, sticky="w")

    def on_run():
        # Gather user selections
        sel_function_fields = [f for f, v in function_vars.items() if v.get()]
        sel_macro_fields = [f for f, v in macro_vars.items() if v.get()]
        sel_variable_fields = [f for f, v in variable_vars.items() if v.get()]
        sel_formats = [f for f, v in format_vars.items() if v.get()]
        outdir = Path(save_dir.get())

        # Ask for the C source file
        cfile = filedialog.askopenfilename(
            title="Select C source file",
            filetypes=[("C files","*.c"), ("All files","*.*")]
        )
        if not cfile:
            return

        # Parse the file
        src = open(cfile, encoding="utf-8").read()
        functions, macros, variables = parse_file(src)
        stem = Path(cfile).stem

        # Precompute the Excel path (used for .xlsx and .docx output)
        xlsx_path = outdir / f"{stem}.xlsx"

        # Excel export
        if "Excel" in sel_formats:
            write_excel(str(xlsx_path), functions, macros, variables, 
                       sel_function_fields, sel_macro_fields, sel_variable_fields)
            open_file(str(xlsx_path))

        # Word export
        if "Word" in sel_formats:
            # write_docx takes the Excel path to derive the .docx alongside it
            write_docx(str(xlsx_path), stem, functions, macros, variables,
                      sel_function_fields, sel_macro_fields, sel_variable_fields)
            open_file(str(xlsx_path.with_suffix('.docx')))

        # Markdown export
        if "MD" in sel_formats:
            md_path = outdir / f"{stem}.md"
            write_markdown(str(md_path), functions, macros, variables,
                          sel_function_fields, sel_macro_fields, sel_variable_fields)
            open_file(str(md_path))

        # Close the GUI
        root.destroy()

    ttk.Button(settings_frame, text="Run", command=on_run).grid(row=2, column=0, pady=15)
    root.mainloop()

def main():
    ap = argparse.ArgumentParser(description="Extract and document Runnables, Macros, and Variables")
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
        functions, macros, variables = parse_file(src)
    except Exception as e:
        print(f"❌ Error parsing {args.file}: {e}", file=sys.stderr)
        sys.exit(1)

    # Output JSON for all parsed data
    output = {
        "functions": functions,
        "macros": macros,
        "variables": variables
    }
    print(json.dumps(output, indent=2))

    swcName = Path(args.file).stem
    write_docx(args.file, swcName, functions, macros, variables, [
        "Line Number","Name","Description","Syntax","Triggers","In-Parameters","Out-Parameters",
        "Return Value","Function Type","Inputs","Outputs",
        "Invoked Operations","Used Data Types","Sync/Async","Reentrancy"
    ], [
        "Line Number", "Name", "Value"
    ], [
        "Line Number", "Name", "Data Type", "Initial Value", "Scope"
    ])

if __name__ == "__main__":
    main()