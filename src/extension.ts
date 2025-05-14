/* 
 - Using python instead of ts for easy debugging for me.
 - Introducing the new methodology by parsing the Runnables and helper functions directly instead of parsing the template comment block.
 - Updating the version by the newly added major changes.
*/

import * as vscode from 'vscode';
import * as path from 'path';
import { execFile } from 'child_process';
import * as fs from 'fs';
import * as Excel from 'exceljs';

const HEADERS = [
  "Name",
  "Return Value",
  "In-Parameters",
  "Out-Parameters",
  "Function Type",
  "Description",
  "Triggers",
  "Inputs",
  "Outputs",
  "Invoked Operations",
  "Used Data Types",
];

export function activate(ctx: vscode.ExtensionContext) {
  ctx.subscriptions.push(
    vscode.commands.registerCommand(
      'runnable-info-extractor.runnable-info-extractor',
      () => {
        const editor = vscode.window.activeTextEditor;
        if (!editor) {
          return vscode.window.showWarningMessage('Open a C file first.');
        }

        const uri = editor.document.uri;
        const srcPath = uri.fsPath;
        const swcName = getSWCName(uri);
        vscode.window.showInformationMessage(`🚀 Extracting for SWC: ${swcName}`);

        const pyPath = path.join(ctx.extensionPath, 'parser.py');
        const candidates =
          process.platform === 'win32'
            ? ['py', 'python', 'python3']
            : ['python3', 'python'];
        let idx = 0;

        function tryExec() {
          if (idx >= candidates.length) {
            return vscode.window.showErrorMessage(
              `No Python interpreter found. Tried: ${candidates.join(', ')}`
            );
          }
          const pythonCmd = candidates[idx++];
          vscode.window.showInformationMessage(`🔍 Running ${pythonCmd}…`);

          execFile(pythonCmd, [pyPath, srcPath], {}, async (err, stdout, stderr) => {
            if (err) {
              if (
                (err as any).code === 'ENOENT' ||
                /not found/i.test(err.message)
              ) {
                return tryExec();
              }
              vscode.window.showErrorMessage(`❌ ${pythonCmd} error: ${err.message}`);
              if (stderr) {
                vscode.window.showWarningMessage(stderr);
              }
              return;
            }
            if (stderr) {
              vscode.window.showWarningMessage(`⚠️ parser stderr:\n${stderr}`);
            }

            let rows: any[];
            try {
              rows = JSON.parse(stdout);
            } catch (je) {
              return vscode.window.showErrorMessage(`❌ JSON parse error: ${je}`);
            }

            vscode.window.showInformationMessage(`✅ Found ${rows.length} functions.`);

            try {
              await writeExcel(srcPath, swcName, rows);
              vscode.window.showInformationMessage(`💾 Saved ${swcName}.xlsx`);
            } catch (we: any) {
              vscode.window.showErrorMessage(`❌ Excel write error: ${we.message || we}`);
            }
          });
        }

        tryExec();
      }
    )
  );
}

function getSWCName(uri: vscode.Uri): string {
  const full = path.basename(uri.fsPath);
  return full.replace(/\.[^.]+$/, '');
}

async function writeExcel(
  srcPath: string,
  swcName: string,
  rows: any[]
) {
  const outDir = path.dirname(srcPath);
  const outPath = path.join(outDir, `${swcName}.xlsx`);

  const workbook = fs.existsSync(outPath)
    ? await new Excel.Workbook().xlsx.readFile(outPath)
    : new Excel.Workbook();

  const ws = workbook.getWorksheet('Runnables') || workbook.addWorksheet('Runnables');

  if (ws.rowCount === 0) {
    ws.addRow(HEADERS);
  }

  // Style header row
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFF0000' } }; // red font
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }, // yellow fill
    };
  });

  // Append data rows
  for (const r of rows) {
    ws.addRow([
      r.name,
      r.ret,
      r.inParams.join(', '),
      r.outParams.join(', '),
      r.fnType,
      '', // Description
      r.trigger,
      r.inputs.join(', '),
      r.outputs.join(', '),
      r.invoked.join(', '),
      r.used.join(', '),
    ]);
  }

  await workbook.xlsx.writeFile(outPath);
}

export function deactivate() {}
