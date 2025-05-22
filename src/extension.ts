/*  Documentation-Slayer VS-Code side
 *  – runs parser.py in GUI mode whenever the command is invoked
 */

import * as vscode from 'vscode';
import * as path   from 'path';
import { execFile } from 'child_process';

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

  const pyCandidates = process.platform === 'win32'
    ? ['py', 'python', 'python3']
    : ['python3', 'python'];

  let pythonExe: string | null = null;
  for (const exe of pyCandidates) {
    try {
      await new Promise<void>((res, rej) => {
        execFile(exe, ['--version'], err => err ? rej(err) : res());
      });
      pythonExe = exe;
      break;
    } catch {
      // try next
    }
  }

  if (!pythonExe) {
    vscode.window.showErrorMessage(
      `No Python interpreter found. Tried: ${pyCandidates.join(', ')}`
    );
    return;
  }

  const pyPath = path.join(ctx.extensionPath, 'parser.py');
  execFile(pythonExe, [pyPath], err => {
    if (err) {
      vscode.window.showErrorMessage(`Failed to launch GUI: ${err.message}`);
    }
  });
  // execFile(pythonExe, [pyPath], …)
}

export function deactivate() {}
