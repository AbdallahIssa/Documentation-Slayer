/*  Documentation-Slayer VS-Code side
 *  – runs parser.py
 *  – writes Excel    (.xlsx)
 *  – writes Markdown (.md)
 *  – one combined notification with Open-Excel / Open-Word / Open-MD
 */

import * as vscode from 'vscode';
import * as path   from 'path';
import { execFile } from 'child_process';
import * as fs     from 'fs';
import * as Excel  from 'exceljs';

const HEADERS = [
  'Name', 'Syntax', 'Return Value', 'In-Parameters', 'Out-Parameters',
  'Function Type', 'Description', 'Sync/Async', 'Reentrancy',
  'Triggers', 'Inputs', 'Outputs',
  'Invoked Operations', 'Used Data Types'
];

export function activate(ctx: vscode.ExtensionContext) {
  ctx.subscriptions.push(
    vscode.commands.registerCommand(
      'Run-Documentation-Slayer.Run-Documentation-Slayer',
      () => runExtractor(ctx)
    )
  );
}

async function runExtractor(ctx: vscode.ExtensionContext) {
  const editor = vscode.window.activeTextEditor;
  if (!editor) {
    vscode.window.showWarningMessage('Open a C file first.');
    return;
  }

  const uri     = editor.document.uri;
  const srcPath = uri.fsPath;
  const swcName = path.basename(srcPath, path.extname(srcPath));

  /* find python */
  const pythonExe = await findPython();
  if (!pythonExe) return;

  /* run parser.py */
  const pyPath = path.join(ctx.extensionPath, 'parser.py');
  const { stdout, stderr } = await runProcess(pythonExe, [pyPath, srcPath]);

  if (stderr.trim()) {
    vscode.window.showWarningMessage(`⚠️ parser stderr:\n${stderr}`);
  }

  let rows: any[];
  try {
    rows = JSON.parse(stdout);
  } catch (e) {
    vscode.window.showErrorMessage(`❌ JSON parse error: ${(e as Error).message}`);
    return;
  }

  /* output paths */
  const outDir     = path.dirname(srcPath);
  const excelPath  = path.join(outDir, `${swcName}.xlsx`);
  const wordPath   = path.join(outDir, `${swcName}.docx`);
  const mdPath     = path.join(outDir,  `${swcName}.md`);

  /* save Excel & Markdown with progress spinner */
  try {
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title:    `Saving ${swcName}.xlsx / .md…`,
        cancellable: false
      },
      async () => {
        await writeExcel(excelPath, rows);
        await writeMarkdown(mdPath, rows);
      }
    );
  } catch (err: any) {
    vscode.window.showErrorMessage(
      `❌ File write error: ${err?.message ?? String(err)}`
    );
    return;
  }

  /* final clickable notification */
  const choice = await vscode.window.showInformationMessage(
    `✅ Found ${rows.length} functions and saved ${swcName}.xlsx`,
    'Open Excel', 'Open Word', 'Open MD'
  );

  if (choice === 'Open Excel') {
    vscode.env.openExternal(vscode.Uri.file(excelPath));
  } else if (choice === 'Open Word') {
    vscode.env.openExternal(vscode.Uri.file(wordPath));
  } else if (choice === 'Open MD') {
    const doc = await vscode.workspace.openTextDocument(mdPath);
    vscode.window.showTextDocument(doc, { preview: false });
  }
}

/* ───────── helpers ───────────────────────────────────────── */

async function findPython(): Promise<string | null> {
  const cands = process.platform === 'win32'
    ? ['py', 'python', 'python3']
    : ['python3', 'python'];

  for (const exe of cands) {
    try { await runProcess(exe, ['--version']); return exe; }
    catch {/* next */}
  }
  vscode.window.showErrorMessage(`No Python interpreter found. Tried: ${cands.join(', ')}`);
  return null;
}

function runProcess(cmd: string, args: string[]) {
  return new Promise<{ stdout: string; stderr: string }>((resolve, reject) => {
    execFile(cmd, args, { encoding: 'utf8', maxBuffer: 10_000_000 },
      (err, stdout, stderr) => err ? reject(err) : resolve({ stdout, stderr }));
  });
}

/* ─── Excel ──────────────────────────────────────────────── */
async function writeExcel(file: string, rows: any[]) {
  const wb = fs.existsSync(file)
    ? await new Excel.Workbook().xlsx.readFile(file)
    : new Excel.Workbook();

  const ws = wb.getWorksheet('Runnables and static functions')
            ?? wb.addWorksheet('Runnables and static functions');

  if (ws.rowCount === 0) {
    ws.addRow(HEADERS);
    const hdr = ws.getRow(1);
    hdr.eachCell(c => {
      c.font = { bold: true, color: { argb:'FFFF0000' } };
      c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFF00' } };
    });
  }

  for (const r of rows) {
    ws.addRow([
      r.name,
      r.syntax,
      r.ret,
      r.inParams.join(', '),
      r.outParams.join(', '),
      r.fnType,
      '',                         // Description
      r.Sync_Async,
      r.Reentrancy,
      r.trigger,
      r.inputs.join(', '),
      r.outputs.join(', '),
      r.invoked.join(', '),
      r.used.join(', ')
    ]);
  }

  ws.columns.forEach(col => {
    let max = 12;
    col.eachCell?.(c => max = Math.max(max, (c.value?.toString().length ?? 0)+2));
    col.width = max;
  });

  await wb.xlsx.writeFile(file);
}

/* ─── Markdown ───────────────────────────────────────────── */
async function writeMarkdown(file: string, rows: any[]) {
  const lines: string[] = [];
  for (const r of rows) {
    lines.push(`## ${r.name}\n`);
    lines.push('| Field | Value |');
    lines.push('|-------|-------|');
    lines.push(`| Syntax | \`${r.syntax}\` |`);
    lines.push(`| Sync/Async | \`${r.Sync_Async}\` |`);
    lines.push(`| Reentrancy | \`${r.Reentrancy}\` |`);
    lines.push(`| Return Value | \`${r.ret}\` |`);
    lines.push(`| In-Parameters | ${r.inParams.join(', ')} |`);
    lines.push(`| Out-Parameters | ${r.outParams.join(', ')} |`);
    lines.push(`| Function Type | ${r.fnType} |`);
    lines.push(`| Triggers | ${r.trigger} |`);
    lines.push(`| Inputs | ${r.inputs.join(', ')} |`);
    lines.push(`| Outputs | ${r.outputs.join(', ')} |`);
    lines.push(`| Invoked Operations | ${r.invoked.join(', ')} |`);
    lines.push(`| Used Data Types | ${r.used.join(', ')} |`);
    lines.push(`| Description | ${""} |`);
    lines.push(''); // blank line between functions
  }
  await fs.promises.writeFile(file, lines.join('\n'), 'utf8');
}

export function deactivate() { }
