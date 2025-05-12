// // The module 'vscode' contains the VS Code extensibility API
// // Import the module and reference it with the alias vscode in your code below
// import * as vscode from 'vscode';

// // This method is called when your extension is activated
// // Your extension is activated the very first time the command is executed
// export function activate(context: vscode.ExtensionContext) {

// 	// Use the console to output diagnostic information (console.log) and errors (console.error)
// 	// This line of code will only be executed once when your extension is activated
// 	console.log('Congratulations, your extension "runnable-info-extractor" is now active!');

// 	// The command has been defined in the package.json file
// 	// Now provide the implementation of the command with registerCommand
// 	// The commandId parameter must match the command field in package.json
// 	const disposable = vscode.commands.registerCommand('runnable-info-extractor.helloWorld', () => {
// 		// The code you place here will be executed every time your command is executed
// 		// Display a message box to the user
// 		vscode.window.showInformationMessage('Hello World from Runnable_Info_Extractor!');
// 	});

// 	context.subscriptions.push(disposable);
// }

// // This method is called when your extension is deactivated
// export function deactivate() {}

/* Abdallah Issa starts here */
// import * as vscode from 'vscode';
// import * as ExcelJS from 'exceljs';

// export function activate(context: vscode.ExtensionContext) {
//     let disposable = vscode.commands.registerCommand('runnable-info-extractor.runnable-info-extractor', async () => {
//         const editor = vscode.window.activeTextEditor;
//         if (!editor) {
//             vscode.window.showErrorMessage('No active editor');
//             return;
//         }

//         const text = editor.document.getText();

//         // Extract function name
//         const functionName = /Runnable Entity Name:\s*(\w+)/.exec(text)?.[1] || '';
//         const trigger = /trigger conditions occurred:\s*\n\s*\*\s*-\s*(.+)/.exec(text)?.[1] || '';

//         const inputs = extractBlock(text, /Input Interfaces:[\s\S]*?Explicit S\/R API:[\s\S]*?[-]+\n([\s\S]*?)\n\s*\*/);
//         const outputs = extractBlock(text, /Output Interfaces:[\s\S]*?Explicit S\/R API:[\s\S]*?[-]+\n([\s\S]*?)\n\s*\*/);
//         const operations = extractBlock(text, /Client\/Server Interfaces:[\s\S]*?Server Invocation:[\s\S]*?[-]+\n([\s\S]*?)\n\s*\*/);

//         // Excel Export
//         const workbook = new ExcelJS.Workbook();
//         const sheet = workbook.addWorksheet('Runnables');
//         sheet.addRow(['Function Name', 'Trigger', 'Inputs', 'Outputs', 'Invoked Operations']);
//         sheet.addRow([functionName, trigger, inputs.join('\n'), outputs.join('\n'), operations.join('\n')]);

//         const uri = await vscode.window.showSaveDialog({ filters: { 'Excel Files': ['xlsx'] } });
//         if (uri) {
//             await workbook.xlsx.writeFile(uri.fsPath);
//             vscode.window.showInformationMessage('Excel exported!');
//         }
//     });

//     context.subscriptions.push(disposable);
// }

// function extractBlock(text: string, regex: RegExp): string[] {
//     const match = regex.exec(text);
//     if (!match) return [];
//     return match[1].split('\n').map(line => line.trim()).filter(line => line);
// }

// // export function deactivate() {}
// import * as vscode from 'vscode';
// import * as Excel from 'exceljs';

// export function activate(context: vscode.ExtensionContext) {
//   let disposable = vscode.commands.registerCommand('runnable-info-extractor.runnable-info-extractor', async () => {
//     const editor = vscode.window.activeTextEditor;
//     if (!editor) {
//       vscode.window.showErrorMessage('No active text editor found.');
//       return;
//     }

//     const document = editor.document;
//     const text = document.getText();

//     // Match every runnable block in the file
//     const runnableRegex = /Runnable Entity Name:\s*(.*)\s*[\s\S]*?trigger conditions occurred:\s*\n\s*\*\s*- (.*)\s*[\s\S]*?Input Interfaces:[\s\S]*?Explicit S\/R API:[\s\S]*?[-]+\n([\s\S]*?)\n\s*\*\n[\s\S]*?Output Interfaces:[\s\S]*?Explicit S\/R API:[\s\S]*?[-]+\n([\s\S]*?)\n\s*\*\n[\s\S]*?Client\/Server Interfaces:[\s\S]*?Server Invocation:[\s\S]*?[-]+\n([\s\S]*?)\n\s*\*/g;

//     const matches = Array.from(text.matchAll(runnableRegex));

//     if (matches.length === 0) {
//       vscode.window.showWarningMessage('No runnable blocks found in the current file.');
//       return;
//     }

//     // Create workbook & sheet
//     const workbook = new Excel.Workbook();
//     const sheet = workbook.addWorksheet('Runnables');

//     // Add header row
//     sheet.addRow(['Function Name', 'Trigger', 'Inputs', 'Outputs', 'Invoked Operations']);

//     // For each matched runnable block, add a row
//     for (const match of matches) {
//       const functionName = match[1].trim();
//       const trigger = match[2].trim();
//       const inputs = match[3].trim();
//       const outputs = match[4].trim();
//       const invokedOps = match[5].trim();

//       // Add 1 row per function (multiline cells are fine in Excel)
//       sheet.addRow([functionName, trigger, inputs, outputs, invokedOps]);
//     }

//     // Ask user where to save Excel file
//     const fileUri = await vscode.window.showSaveDialog({
//       filters: { 'Excel Files': ['xlsx'] },
//       defaultUri: vscode.Uri.file('runnables.xlsx')
//     });

//     if (fileUri) {
//       await workbook.xlsx.writeFile(fileUri.fsPath);
//       vscode.window.showInformationMessage('Excel file saved successfully!');
//     } else {
//       vscode.window.showInformationMessage('Save cancelled.');
//     }
//   });

//   context.subscriptions.push(disposable);
// }

// export function deactivate() {}

/* Rayes code */
/* ────────────────────────────────────────────────────────────────────────────────
   VS‑Code command: “runnableDoc.exportExcel”
   ────────────────────────────────────────────────────────────────────────────── */
import * as vscode from 'vscode';
import * as path   from 'path';
import * as Excel  from 'exceljs';          //  npm i exceljs

interface RunnableInfo {
  name:        string;
  trigger:     string;
  inputs:      string[];
  outputs:     string[];
  invokedOps:  string[];
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

      const defaultFile = path.join(
        path.dirname(editor.document.uri.fsPath), 'runnables.xlsx'
      );
      const uri = await vscode.window.showSaveDialog({
        defaultUri: vscode.Uri.file(defaultFile),
        filters: { Excel: ['xlsx'] }
      });
      if (!uri) { return; }

      await writeExcel(uri.fsPath, runnables);
      vscode.window.showInformationMessage(
        `Exported ${runnables.length} runnables → ${uri.fsPath}`
      );
    });

  ctx.subscriptions.push(disposable);
}

export function deactivate () { /* nothing */ }

/* ───────────────────────────────── Parsing ──────────────────────────────────── */
function extractRunnables (source: string): RunnableInfo[] {
  /* grab each whole documentation block first */
  const blocks = source.match(
    /\/\*{5}[\s\S]*?End of documentation area[\s\S]*?\*\//g
  ) || [];

  return blocks
    .map(parseBlock)
    .filter((x): x is RunnableInfo => Boolean(x));
}

function parseBlock(block: string): RunnableInfo | null {
  /* skip dummy prototype blocks outright */
  if (/Runnable prototype:/i.test(block)) return null;

  /* strip leading " * " for uniform text */
  const lines = block
    .split('\n')
    .map(l => l.replace(/^\s*\*\s?/, '').trimEnd());

  /* function / runnable name */
  const nameLine = lines.find(l => /Runnable Entity Name:/i.test(l));
  const name = nameLine?.split(':').pop()?.trim();
  if (!name) return null;

  /* trigger = the very next non‑blank line after the heading */
  const trigIdx = lines.findIndex(l =>
    /trigger conditions occurred/i.test(l)
  );
  const trigger =
    trigIdx > -1 && lines[trigIdx + 1] ? lines[trigIdx + 1].trim() : '';

  /* Inputs, Outputs, Invoked Operations */
  const inputs  = collectSection(lines, 'Input Interfaces',  'Inter Runnable');
  const outputs = collectSection(lines, 'Output Interfaces', 'Inter Runnable');
  const invoked = collectSubSection(lines, 'Server Invocation');

  return { name, trigger, inputs, outputs, invokedOps: invoked };
}

// function parseBlock (block: string): RunnableInfo | null {
//   /* strip leading “ * ” for easier scanning */
//   const lines = block
//     .split('\n')
//     .map(l => l.replace(/^\s*\*\s?/, '').trimEnd());

//   const getAfter = (
//     startsWith: string,
//     stop: (l: string) => boolean = l => !l
//   ): string[] => {
//     const i = lines.findIndex(l => l.startsWith(startsWith));
//     if (i === -1) return [];
//     const out: string[] = [];
//     for (let j = i + 1; j < lines.length && !stop(lines[j]); j++) {
//       const ln = lines[j];
//       if (ln.startsWith('-----------------')) continue;   // visual divider
//       out.push(ln.trim());
//     }
//     return out.filter(Boolean);
//   };

//   const name    = /Runnable Entity Name:\s*([A-Za-z0-9_]+)/.exec(block)?.[1];
//   if (!name) return null;

//   const trigger = getAfter(
//     'Executed if at least one of the following trigger conditions occurred:'
//   )[0] || '';

//   const inputs  = getAfter(
//     'Explicit S/R API:',
//     l => l.startsWith('Output Interfaces') ||
//          l.startsWith('Inter Runnable')   ||
//          l.startsWith('Client/Server')
//   );

//   const outputs = getAfter(
//     'Output Interfaces:',
//     l => l.startsWith('Inter Runnable') ||
//          l.startsWith('Client/Server')
//   );

//   const invoked = getAfter(
//     'Server Invocation:',
//     l => l.startsWith('*') || l.startsWith('/') || !l
//   );

//   return { name, trigger, inputs, outputs, invokedOps: invoked };
// }



/* ─────────────────────────── Excel writer ──────────────────────────────────── */
async function writeExcel (file: string, data: RunnableInfo[]) {
  const wb = new Excel.Workbook();
  const ws = wb.addWorksheet('Runnables');

  ws.addRow(['Function Name', 'Trigger', 'Inputs', 'Outputs', 'Invoked Operations']);

  for (const r of data) {
    ws.addRow([
      r.name,
      r.trigger,
      r.inputs.join('\n'),
      r.outputs.join('\n'),
      r.invokedOps.join('\n')
    ]);
  }
  /* auto‑fit columns */
  ws.columns.forEach(col => {
    let max = 10;
    /* auto‑fit columns */
  for (const col of ws.columns) {
    if (!col) continue;               // skip undefined / null slots

    let max = 10;
    col.eachCell?.(cell => {          // “?.“ = only if eachCell exists
      max = Math.max(max, (cell.value?.toString().length ?? 0) + 2);
    });

    col.width = max;
  }
    // col.eachCell(c => { max = Math.max(max, (c.value?.toString().length ?? 0) + 2); });
    // col.width = max;
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

    /* stop when we reach the next major heading */
    if (ln.includes(stopLabel) || ln.includes('Client/Server Interfaces')) break;

    /* skip purely decorative or heading lines */
    if (/^[-=]+$/.test(ln)) continue;                // "-----"  or  "====="
    if (/^Explicit S\/R API:?/i.test(ln)) continue;  // "Explicit S/R API:"
    if (!ln) continue;                               // blank line

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

    /* stop at the end of the sub‑section */
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