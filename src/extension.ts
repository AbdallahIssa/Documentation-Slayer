/* ────────────────────────────────────────────────────────────────────────────────
   VS-Code command: “runnable-info-extractor.runnable-info-extractor”
   ────────────────────────────────────────────────────────────────────────────── */
// import * as vscode from 'vscode';
// import * as path   from 'path';
// import * as Excel  from 'exceljs';          //  npm i exceljs

// interface RunnableInfo {
//   name:       string;
//   triggers:   string[];     // now an array
//   inputs:     string[];
//   outputs:    string[];
//   invokedOps: string[];
// }

// /* ─────────────── Extension entry point ─────────────── */
// export function activate (ctx: vscode.ExtensionContext) {
//   const disposable = vscode.commands.registerCommand(
//     'runnable-info-extractor.runnable-info-extractor',
//     async () => {
//       const editor = vscode.window.activeTextEditor;
//       if (!editor) { return vscode.window.showErrorMessage('Open a C file first'); }

//       const runnables = extractRunnables(editor.document.getText());
//       if (!runnables.length) {
//         return vscode.window.showWarningMessage('No runnable documentation blocks found.');
//       }

//       const swcName    = getSwcName(editor.document.uri); // vscode.window.showInputBox();
//       const defaultFile = path.join(
//         path.dirname(editor.document.uri.fsPath),
//         `${swcName}.xlsx`
//       );
//       const uri = await vscode.window.showSaveDialog({
//         defaultUri: vscode.Uri.file(defaultFile),
//         filters: { Excel: ['xlsx'] }
//       });
//       if (!uri) { return; }

//       await writeExcel(uri.fsPath, runnables, swcName);
//       vscode.window.showInformationMessage(
//         `Exported ${runnables.length} runnables → ${uri.fsPath}`
//       );
//     });

//   ctx.subscriptions.push(disposable);
// }

// export function deactivate () { /* nothing */ }

// /* ───────────────────────────────── Parsing ──────────────────────────────────── */
// function extractRunnables (source: string): RunnableInfo[] {
  
// const blocks = source.match(/\/\*{2,}[\s\S]*?\*\//g) || [];

//   return blocks
//     .map(parseBlock)
//     .filter((x): x is RunnableInfo => Boolean(x));
// }

// function parseBlock(block: string): RunnableInfo | null {
//    /* if it doesn’t even contain “Runnable Entity Name:” it isn’t a real runnable */
//    if (!/Runnable Entity Name:/i.test(block)) return null;

//   /* strip leading " * " for uniform text */
//   const lines = block
//     .split('\n')
//     .map(l => l.replace(/^\s*\*\s?/, '').trimEnd());

//   /* function / runnable name */
//   const nameLine = lines.find(l => /Runnable Entity Name:/i.test(l));
//   const name = nameLine?.split(':').pop()?.trim();
//   if (!name) return null;

//   /* full trigger list */
//   const triggers = collectTriggers(lines);

//   /* Inputs, Outputs, Invoked Operations */
//   const inputs  = collectSection(lines, 'Input Interfaces',  'Inter Runnable');
//   const outputs = collectSection(lines, 'Output Interfaces', 'Inter Runnable');
//   const invoked = collectSubSection(lines, 'Server Invocation');

//   return { name, triggers, inputs, outputs, invokedOps: invoked };
// }

// /* ─────────────────────────── Excel writer ──────────────────────────────────── */
// async function writeExcel (file: string,
//                            data: RunnableInfo[],
//                            sheetName: string) {
//   const wb = new Excel.Workbook();
//   const ws = wb.addWorksheet(sheetName);

//   /* ── header ───────────────────────────── */
//   ws.addRow(['Function Name', 'Trigger(s)',
//              'Inputs', 'Outputs', 'Invoked Operations']);

//   const header = ws.getRow(1);
//   header.font = { bold: true, color: { argb: 'FFFF0000' } };   // red
//   header.fill = {
//     type: 'pattern',
//     pattern: 'solid',
//     fgColor: { argb: 'FFFFFF00' }                              // yellow
//   };
//   header.commit();

//   /* ── data rows ────────────────────────── */
//   for (const r of data) {
//     ws.addRow([
//       r.name,
//       r.triggers.join('\n'),
//       r.inputs.join('\n'),
//       r.outputs.join('\n'),
//       r.invokedOps.join('\n')
//     ]);
//   }

//   /* auto-fit columns (keeps earlier logic) */
//   ws.columns.forEach(col => {
//     if (!col) return;
//     let max = 10;
//     col.eachCell?.(cell =>
//       max = Math.max(max,
//                      (cell.value?.toString().length ?? 0) + 2));
//     col.width = max;
//   });

//   await wb.xlsx.writeFile(file);
// }


// /* ──────────────────────────── parsing helpers ──────────────────────────── */
// function collectSection(
//   lines: string[],
//   startLabel: string,
//   stopLabel: string
// ): string[] {
//   const start = lines.findIndex(l => l.includes(startLabel));
//   if (start === -1) return [];

//   const result: string[] = [];
//   for (let i = start + 1; i < lines.length; i++) {
//     const ln = lines[i].trim();

//     /* stop when we hit the next major heading */
//     if (ln.includes(stopLabel) ||
//         ln.includes('Client/Server Interfaces') ||
//         ln.includes('Output Interfaces')) break;

//     /* skip decorative / unwanted lines */
//     if (/^[-=]+$/.test(ln)) continue;
//     if (/^Explicit S\/R API:?/i.test(ln)) continue;
//     if (/^Implicit S\/R API:?/i.test(ln)) continue;
//     if (/DO NOT CHANGE THIS COMMENT!/i.test(ln)) continue;
//     if (/<<\s*(Start|End) of documentation area\s*>>/i.test(ln)) continue;
//     if (!ln) continue;

//     result.push(ln);
//   }
//   return result;
// }

// function collectSubSection(lines: string[], header: string): string[] {
//   const idx = lines.findIndex(l => l.includes(header));
//   if (idx === -1) return [];

//   const result: string[] = [];
//   for (let i = idx + 1; i < lines.length; i++) {
//     const ln = lines[i].trim();

//     /* stop at the end of the sub-section */
//     if (ln.startsWith('*') || ln.startsWith('/') || !ln) break;

//     /* ignore explanatory lines */
//     if (/^Synchronous/i.test(ln) ||
//         /^Returned Application/i.test(ln) ||
//         /^Technical Application/i.test(ln)) continue;

//     /* keep only real call prototypes (must look like “func(…)”) */
//     if (!ln.includes('(')) continue;

//     result.push(ln);
//   }
//   return result;
// }

// function collectTriggers(lines: string[]): string[] {
//   const idx = lines.findIndex(l => /trigger conditions occurred/i.test(l));
//   if (idx === -1) return [];

//   const result: string[] = [];
//   for (let i = idx + 1; i < lines.length; i++) {
//     const ln = lines[i].trim();

//     /* a blank line or a heading ends the section */
//     // if (!ln || /Interfaces|Runnable Entity Name:/i.test(ln)) break;
//     /* stop only at the **next heading** (look for lines ending with “:”) */
//     if (!ln) break;                       // blank line = end of block
//     if (/^[A-Z].*?:\s*$/.test(ln)) break; // e.g. “Input Interfaces:”
//     if (/^[-=]+$/.test(ln)) continue;     // decorative “-----”
//     result.push(ln);
//   }
//   return result;
// }

// function getSwcName(uri: vscode.Uri): string {
//   const full = path.basename(uri.fsPath);
//   return full.replace(/\.[^.]+$/, '');
// }

/* The new methodology */

// import * as vscode from "vscode";
// import * as path from "path";
// import * as fs from "fs";
// import * as Excel from "exceljs";

// /* ──────────────────────────────── 1.  Regexes ──────────────────────────────── */
// const REGEX = {
//   /* comment block immediately preceding a function
//      (captures the full block + the runnable trigger line)                                 */
//   // commentBlock: /\/\*[\s\S]*?\*\/,

//   /* “triggered on …” inside the comment block                                             */
//   trigger: /-\s*triggered\s+on\s+([^\n\r]+)/,

//   /* complete function – grabs signature + body (non-greedy)                              */
//   // functionBlock:
//   //   /(?:static\s+)?(?:FUNC\s*\([^)]*\)\s*)?([A-Za-z_][A-Za-z0-9_]*)\s*\(([^)]*)\)\s*\{([\s\S]*?)^\}/gm,

//   /* return type in a FUNC(…) macro              e.g.  FUNC(void, …)                      */
//   funcReturn: /FUNC\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)/,

//   /* parameters list splitter (very forgiving)                                            */
//   splitParams: /,(?![^\(\)]*\))/,

//   /* strip parameter qualifiers (const, volatile, etc.)                                   */
//   paramName: /\b([A-Za-z_][A-Za-z0-9_]*)\s*$/,

//   /* Inputs / Outputs / Calls inside the body                                             */
//   rteRead: /\bRte_Read_[A-Za-z0-9_]+\s*\(/g,
//   rteWrite: /\bRte_(?:Write|IrvWrite)_[A-Za-z0-9_]+\s*\(/g,
//   rteCall: /\bRte_Call_[A-Za-z0-9_]+\s*\(/g,
//   plainCall: /\b([A-Za-z_][A-Za-z0-9_]*)\s*\(/g,

//   /* local used data types                                                                */
//   usedTypes: /\b([A-Za-z_][A-Za-z0-9_]*)\s+[A-Za-z_][A-Za-z0-9_]*\s*(?:[=;])/g,
// };

// /* ─────────────────────────────── 2.  Excel helpers ─────────────────────────── */
// const HEADERS = [
//   "Name",
//   "Return Value",
//   "In-Parameters",
//   "Out-Parameters",
//   "Function Type",
//   "Description",
//   "Trigger",
//   "Inputs",
//   "Outputs",
//   "Invoked Operations",
//   "Used Data Types",
// ];

// /* ─────────────────────────────── 3.  Activate cmd ──────────────────────────── */
// export function activate(ctx: vscode.ExtensionContext) {
//   ctx.subscriptions.push(
//     vscode.commands.registerCommand(
//       'runnable-info-extractor.runnable-info-extractor',
//       async () => {
//         const editor = vscode.window.activeTextEditor;
//         if (!editor) {
//           return vscode.window.showWarningMessage("No active editor.");
//         }

//         const text = editor.document.getText();
//         const rows = parseFile(text);

//         if (!rows.length) {
//           return vscode.window.showInformationMessage("No runnables/static functions found.");
//         }

//         const excelPath = path.join(
//           path.dirname(editor.document.uri.fsPath),
//           `${path.basename(editor.document.uri.fsPath)}-runnables.xlsx`
//         );
//         await writeExcel(excelPath, rows);
//         vscode.window.showInformationMessage(`Runnable info written to ${excelPath}`);
//       }
//     )
//   );
// }

// /* ─────────────────────────────── 4.  Core parser ──────────────────────────── */
// interface Row {
//   name: string;
//   ret: string;
//   inParams: string[];
//   outParams: string[];
//   fnType: "Runnable" | "Static";
//   trigger: string;
//   inputs: string[];
//   outputs: string[];
//   invoked: string[];
//   used: string[];
// }

// /* ───────────────────────── FUNCTION SCANNER ──────────────────────── */
// function parseFile(src: string): Row[] {
//   const rows: Row[] = [];
//   const reserved = new Set(["if","for","while","switch","do","else","case"]);
//   /* map of  comment-end-index  →  comment-text  */
//   const commentMap = new Map<number, string>();
//   for (const m of src.matchAll(/\/\*[\s\S]*?\*\//g)) {
//     commentMap.set(m.index! + m[0].length, m[0]);
//   }

//   /* header regex only – body is collected later by brace-counting            *
//  *   · skips control keywords  (if / for / while / switch / do / else / case)
//  *   · allows any return-type tokens (void, uint32, const char*, …)          *
//  *   · tolerates whitespace or comments  before the opening brace       */

// const headerRx =
//   /^[ \t]*(?:static\s+)?(?:FUNC\s*\([^)]*\)\s*)?(?!if\b|for\b|while\b|do\b|switch\b|case\b|else\b)(?:[A-Za-z_]\w*\s+|\*+)*?([A-Za-z_]\w*)\s*\(([^)]*)\)\s*(?:\/\*[\s\S]*?\*\/\s*)*\{/gm;

//   /* nearest comment block (for Trigger)                                    */
//   // const comment = getNearestComment(commentMap, hdr.index!);
//   // const triggerMatch = /-\s*triggered\s+(?:on|by)\s+([^\n\r]+)/im.exec(comment || "");
//   // const trigger = triggerMatch ? triggerMatch[1].trim() : "";

//   for (const hdr of src.matchAll(headerRx)) 
//   {
//     let [full, name, params] = hdr as RegExpMatchArray & { index: number };

//     /* skip control keywords mistakenly seen as “name”                        */
//     if (reserved.has(name)) continue;

//     const bodyStart = hdr.index! + full.length - 1; // at the first “{”
//     let bodyEnd = bodyStart + 1,
//       brace = 1;

//     /* cheap & cheerful brace counter                                         */
//     while (brace && bodyEnd < src.length) {
//       const ch = src[bodyEnd++];
//       if (ch === "{") brace++;
//       else if (ch === "}") brace--;
//     }
//     const body = src.slice(bodyStart + 1, bodyEnd - 1);

//     /* nearest comment block (for Trigger)                                    */
//     const comment = getNearestComment(commentMap, hdr.index!);
//     const triggerMatch = /-\s*triggered\s+(?:on|by)\s+([^\n\r]+)/im.exec(comment || "");
//     const trigger = triggerMatch ? triggerMatch[1].trim() : "";

//     /* return type     tamam                                                       */
//     const ret = 
//       /FUNC\((\w*)/.exec(full)?.[1] || (/static\s+([A-Za-z_][A-Za-z0-9_]*)\s+[A-Za-z_][A-Za-z0-9_]*\(/.exec(full)?.[1] ?? "void");
//     const fnType = full.trimStart().startsWith("static") ? "Static" : "Runnable";

//     /* parameters → names                                                     */
//     const paramArr = params
//       .split(/,(?![^(]*\))/)
//       .map((p) => p.trim())
//       .filter(Boolean);
//     const names = paramArr.map((p) => /\b([A-Za-z_]\w*)\s*$/.exec(p)?.[1] || p);

//     /* classify directions                                                    */
//     const dirs = classifyParams(body, names);
//     const inP = names.filter((p) => dirs[p] === "IN" || dirs[p] === "INOUT");
//     const outP = names.filter((p) => dirs[p] === "OUT" || dirs[p] === "INOUT");

//     /* Inputs / Outputs / Calls                                               */
//     const inputs = unique(matchAll(body, /\bRte_Read_\w+\s*\(/g));
//     const outputs = unique(matchAll(body, /\bRte_(?:Write|IrvWrite)_\w+\s*\(/g));
//     const calls = unique(matchAll(body, /\bRte_Call_\w+\s*\(/g));
//     const locals = unique(
//       matchAll(body, /\b([A-Za-z_]\w*)\s*\(/g).filter(
//         (c) =>
//           !c.startsWith("Rte_") &&
//           !reserved.has(c) &&
//           c !== name
//       )
//     );
//     const invoked = unique([...calls, ...locals]);

//     /* local types                                                            */
//     const used = unique(
//       [...body.matchAll(/\b([A-Za-z_]\w*)\s+[A-Za-z_]\w*\s*(?:[=;])/g)]
//         .map((m) => m[1])
//         .filter(Boolean)
//     );

//     rows.push({
//       name,
//       ret,
//       inParams: inP,
//       outParams: outP,
//       fnType,
//       trigger,
//       inputs,
//       outputs,
//       invoked,
//       used,
//     });
//   }
//   return rows;
// }

// /* helper: extract names from regex matches                                   */
// function matchAll(text: string, rx: RegExp): string[] {
//   const res: string[] = [];
//   for (const m of text.matchAll(rx)) res.push(m[0].replace(/\s*\(/, ""));
//   return res;
// }


// /* ─────────────────────────────── 5.  Helpers ──────────────────────────────── */
// function getNearestComment(map: Map<number, string>, pos: number): string | undefined {
//   const keys = [...map.keys()].filter((k) => k <= pos).sort((a, b) => b - a);
//   return keys.length ? map.get(keys[0]) : undefined;
// }

// function matchAllNames(body: string, rx: RegExp): string[] {
//   const names: string[] = [];
//   for (const m of body.matchAll(rx)) {
//     const token = m[0].replace(/\s*\(/, "");
//     names.push(token);
//   }
//   return names;
// }

// function unique<T>(arr: T[]): T[] {
//   return [...new Set(arr)];
// }

// /* ─────────────────────── 6.  Param direction heuristic ────────────────────── */
// /**
//  * classifyParams()  –  unchanged except for export-removal (internal use)
//  * Very fast heuristic for AUTOSAR-style code.
//  */
// function classifyParams(
//   body: string,
//   params: string[]
// ): Record<string, "IN" | "OUT" | "INOUT"> {
//   const result: Record<string, "IN" | "OUT" | "INOUT"> = {};

//   for (const p of params) {
//     const ptrWrite = new RegExp(String.raw`[\*\(]\s*${p}\s*\)?\s*(?:[+\-*/]?=|[+\-]{2})`, "s");
//     const arrowWrite = new RegExp(String.raw`\b${p}\s*->\s*\w+\s*=`, "s");
//     const incDec = new RegExp(
//       String.raw`(?:\+\+|--)\s*${p}\b|\b${p}\s*(?:\+\+|--)`,
//       "s"
//     );
//     const writeApi = new RegExp(String.raw`\b\w*(?:Write|Set)\w*\s*\([^;]*\b${p}\b`, "s");

//     const isWritten =
//       ptrWrite.test(body) || arrowWrite.test(body) || incDec.test(body) || writeApi.test(body);
//     const isRead = new RegExp(String.raw`\b${p}\b`, "s").test(body);

//     result[p] = isWritten ? (isRead ? "INOUT" : "OUT") : "IN";
//   }
//   return result;
// }

// /* ─────────────────────────────── 7.  Excel writer ─────────────────────────── */
// async function writeExcel(filePath: string, rows: Row[]) {
//   const workbook = fs.existsSync(filePath) ? await new Excel.Workbook().xlsx.readFile(filePath) : new Excel.Workbook();
//   const ws = workbook.getWorksheet("Runnables") || workbook.addWorksheet("Runnables");

//   if (ws.rowCount === 0) ws.addRow(HEADERS);

//   for (const r of rows) {
//     ws.addRow([
//       r.name,
//       r.ret,
//       r.inParams.join(", "),
//       r.outParams.join(", "),
//       r.fnType,
//       "", // description intentionally left blank
//       r.trigger,
//       r.inputs.join(", "),
//       r.outputs.join(", "),
//       r.invoked.join(", "),
//       r.used.join(", "),
//     ]);
//   }

//   await workbook.xlsx.writeFile(filePath);
// }

// /* ─────────────────────────────── 8.  Deactivate ───────────────────────────── */
// export function deactivate() {}





/* 
 - Using python instead of ts for easy debugging for me
 - Introducing the new methodology by parsing the Runnables and helper functions directly instead of parsing the template comment block.
 - Updating the version by the newly added major changes.
*/

import * as vscode from "vscode";
import * as path from "path";
import { execFile } from "child_process";
import * as fs from "fs";
import * as Excel from "exceljs";

const HEADERS = [
  "Name", "Return Value", "In-Parameters", "Out-Parameters", "Function Type",
  "Description", "Trigger", "Inputs", "Outputs", "Invoked Operations", "Used Data Types"
];

export function activate(ctx: vscode.ExtensionContext) {
  ctx.subscriptions.push(
    vscode.commands.registerCommand(
      "runnable-info-extractor.runnable-info-extractor",
      () => {
        vscode.window.showInformationMessage("🚀 Extraction started…");

        const editor = vscode.window.activeTextEditor;
        if (!editor) {
          vscode.window.showWarningMessage("No active editor to extract from.");
          return;
        }

        const srcPath = editor.document.uri.fsPath;
        const pyPath = path.join(ctx.extensionPath, "parser.py");

        execFile(
          "python",
          [pyPath, srcPath],
          { timeout: 30_000 }, // DEBUGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG: wait hena for 30 seconds 34an a3rf by stuck feen
          async (err, stdout, stderr) => {
            if (err) {
              vscode.window.showErrorMessage(
                `❌ Parser failed: ${err.message}`
              );
              return;
            }
            if (stderr) {
              vscode.window.showWarningMessage(
                `⚠️ Parser stderr:\n${stderr}`
              );
            }
            if (!stdout) {
              vscode.window.showWarningMessage(
                "⚠️ Parser returned no output."
              );
              return;
            }

            let rows: any[];
            try {
              rows = JSON.parse(stdout);
            } catch (je) {
              vscode.window.showErrorMessage(
                `❌ JSON parse error: ${je}`
              );
              return;
            }

            vscode.window.showInformationMessage(
              `✅ Parser found ${rows.length} functions.`
            );

            try {
              await writeExcel(srcPath, rows);
              vscode.window.showInformationMessage(
                `💾 Excel written to ${srcPath}-runnables.xlsx`
              );
            } catch (err: any) {
              // Normalize the error message if it's not an Error instance
              const msg = err instanceof Error ? err.message : String(err);
              vscode.window.showErrorMessage(
                `❌ Excel write error: ${msg}`
              );
            }
          }
        );
      }
    )
  );
}

async function writeExcel(filePath: string, rows: any[]) {
  const outPath = `${filePath}-runnables.xlsx`;
  const wb = fs.existsSync(outPath)
    ? await new Excel.Workbook().xlsx.readFile(outPath)
    : new Excel.Workbook();

  const ws = wb.getWorksheet("Runnables") || wb.addWorksheet("Runnables");
  if (ws.rowCount === 0) ws.addRow(HEADERS);

  for (const r of rows) {
    ws.addRow([
      r.name,
      r.ret,
      r.inParams.join(", "),
      r.outParams.join(", "),
      r.fnType,
      "", // description
      r.trigger,
      r.inputs.join(", "),
      r.outputs.join(", "),
      r.invoked.join(", "),
      r.used.join(", "),
    ]);
  }

  await wb.xlsx.writeFile(outPath);
}

export function deactivate() {}
