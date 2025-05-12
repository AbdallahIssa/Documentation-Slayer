/* ────────────────────────────────────────────────────────────────────────────────
   VS-Code command: “runnable-info-extractor.runnable-info-extractor”
   ────────────────────────────────────────────────────────────────────────────── */
import * as vscode from 'vscode';
import * as path   from 'path';
import * as Excel  from 'exceljs';          //  npm i exceljs

interface RunnableInfo {
  name:       string;
  triggers:   string[];     // now an array
  inputs:     string[];
  outputs:    string[];
  invokedOps: string[];
}

/* ─────────────── Extension entry point ─────────────── */
export function activate (ctx: vscode.ExtensionContext) {
  const disposable = vscode.commands.registerCommand(
    'runnable-info-extractor.runnable-info-extractor',
    async () => {
      const editor = vscode.window.activeTextEditor;
      if (!editor) { return vscode.window.showErrorMessage('Open a C file first'); }

      const runnables = extractRunnables(editor.document.getText());
      if (!runnables.length) {
        return vscode.window.showWarningMessage('No runnable documentation blocks found.');
      }

      const swcName    = getSwcName(editor.document.uri); // vscode.window.showInputBox();
      const defaultFile = path.join(
        path.dirname(editor.document.uri.fsPath),
        `${swcName}.xlsx`
      );
      const uri = await vscode.window.showSaveDialog({
        defaultUri: vscode.Uri.file(defaultFile),
        filters: { Excel: ['xlsx'] }
      });
      if (!uri) { return; }

      await writeExcel(uri.fsPath, runnables, swcName);
      vscode.window.showInformationMessage(
        `Exported ${runnables.length} runnables → ${uri.fsPath}`
      );
    });

  ctx.subscriptions.push(disposable);
}

export function deactivate () { /* nothing */ }

/* ───────────────────────────────── Parsing ──────────────────────────────────── */
function extractRunnables (source: string): RunnableInfo[] {
  
const blocks = source.match(/\/\*{2,}[\s\S]*?\*\//g) || [];

  return blocks
    .map(parseBlock)
    .filter((x): x is RunnableInfo => Boolean(x));
}

function parseBlock(block: string): RunnableInfo | null {
   /* if it doesn’t even contain “Runnable Entity Name:” it isn’t a real runnable */
   if (!/Runnable Entity Name:/i.test(block)) return null;

  /* strip leading " * " for uniform text */
  const lines = block
    .split('\n')
    .map(l => l.replace(/^\s*\*\s?/, '').trimEnd());

  /* function / runnable name */
  const nameLine = lines.find(l => /Runnable Entity Name:/i.test(l));
  const name = nameLine?.split(':').pop()?.trim();
  if (!name) return null;

  /* full trigger list */
  const triggers = collectTriggers(lines);

  /* Inputs, Outputs, Invoked Operations */
  const inputs  = collectSection(lines, 'Input Interfaces',  'Inter Runnable');
  const outputs = collectSection(lines, 'Output Interfaces', 'Inter Runnable');
  const invoked = collectSubSection(lines, 'Server Invocation');

  return { name, triggers, inputs, outputs, invokedOps: invoked };
}

/* ─────────────────────────── Excel writer ──────────────────────────────────── */
async function writeExcel (file: string,
                           data: RunnableInfo[],
                           sheetName: string) {
  const wb = new Excel.Workbook();
  const ws = wb.addWorksheet(sheetName);

  /* ── header ───────────────────────────── */
  ws.addRow(['Function Name', 'Trigger(s)',
             'Inputs', 'Outputs', 'Invoked Operations']);

  const header = ws.getRow(1);
  header.font = { bold: true, color: { argb: 'FFFF0000' } };   // red
  header.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' }                              // yellow
  };
  header.commit();

  /* ── data rows ────────────────────────── */
  for (const r of data) {
    ws.addRow([
      r.name,
      r.triggers.join('\n'),
      r.inputs.join('\n'),
      r.outputs.join('\n'),
      r.invokedOps.join('\n')
    ]);
  }

  /* auto-fit columns (keeps earlier logic) */
  ws.columns.forEach(col => {
    if (!col) return;
    let max = 10;
    col.eachCell?.(cell =>
      max = Math.max(max,
                     (cell.value?.toString().length ?? 0) + 2));
    col.width = max;
  });

  await wb.xlsx.writeFile(file);
}


/* ──────────────────────────── parsing helpers ──────────────────────────── */
function collectSection(
  lines: string[],
  startLabel: string,
  stopLabel: string
): string[] {
  const start = lines.findIndex(l => l.includes(startLabel));
  if (start === -1) return [];

  const result: string[] = [];
  for (let i = start + 1; i < lines.length; i++) {
    const ln = lines[i].trim();

    /* stop when we hit the next major heading */
    if (ln.includes(stopLabel) ||
        ln.includes('Client/Server Interfaces') ||
        ln.includes('Output Interfaces')) break;

    /* skip decorative / unwanted lines */
    if (/^[-=]+$/.test(ln)) continue;
    if (/^Explicit S\/R API:?/i.test(ln)) continue;
    if (/^Implicit S\/R API:?/i.test(ln)) continue;
    if (/DO NOT CHANGE THIS COMMENT!/i.test(ln)) continue;
    if (/<<\s*(Start|End) of documentation area\s*>>/i.test(ln)) continue;
    if (!ln) continue;

    result.push(ln);
  }
  return result;
}

function collectSubSection(lines: string[], header: string): string[] {
  const idx = lines.findIndex(l => l.includes(header));
  if (idx === -1) return [];

  const result: string[] = [];
  for (let i = idx + 1; i < lines.length; i++) {
    const ln = lines[i].trim();

    /* stop at the end of the sub-section */
    if (ln.startsWith('*') || ln.startsWith('/') || !ln) break;

    /* ignore explanatory lines */
    if (/^Synchronous/i.test(ln) ||
        /^Returned Application/i.test(ln) ||
        /^Technical Application/i.test(ln)) continue;

    /* keep only real call prototypes (must look like “func(…)”) */
    if (!ln.includes('(')) continue;

    result.push(ln);
  }
  return result;
}

function collectTriggers(lines: string[]): string[] {
  const idx = lines.findIndex(l => /trigger conditions occurred/i.test(l));
  if (idx === -1) return [];

  const result: string[] = [];
  for (let i = idx + 1; i < lines.length; i++) {
    const ln = lines[i].trim();

    /* a blank line or a heading ends the section */
    // if (!ln || /Interfaces|Runnable Entity Name:/i.test(ln)) break;
    /* stop only at the **next heading** (look for lines ending with “:”) */
    if (!ln) break;                       // blank line = end of block
    if (/^[A-Z].*?:\s*$/.test(ln)) break; // e.g. “Input Interfaces:”
    if (/^[-=]+$/.test(ln)) continue;     // decorative “-----”
    result.push(ln);
  }
  return result;
}

function getSwcName(uri: vscode.Uri): string {
  const full = path.basename(uri.fsPath);
  return full.replace(/\.[^.]+$/, '');
}