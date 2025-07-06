import * as vscode from 'vscode';
import { execFile } from 'child_process';

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(
    vscode.commands.registerCommand(
      'Run-Documentation-Slayer.Run-Documentation-Slayer',
      () => runExtractor(context)
    )
  );
}

function runExtractor(context: vscode.ExtensionContext) {
  // Check there’s an active editor (optional, GUI will re-ask anyway)
  if (!vscode.window.activeTextEditor) {
    vscode.window.showWarningMessage('Open a C file first.');
    return;
  }

  // Pick the right executable name per platform
  const exeName = process.platform === 'win32' ? 'Doc-Slayer.exe' : 'Doc-Slayer';

  // Build a URI for the exe inside the extension
  const exeUri = vscode.Uri.joinPath(context.extensionUri, exeName);
  const exePath = exeUri.fsPath;

  // Launch the GUI (no args → interactive mode)
  execFile(exePath, [], (err, stdout, stderr) => {
    if (err) {
      if ((err as any).code === 'ENOENT') {
        vscode.window.showErrorMessage(
          `Cannot find bundled Doc-Slayer.exe at:\n${exePath}\n\n` +
          `Make sure you’ve copied ${exeName} into the extension’s root folder.`
        );
      } else {
        vscode.window.showErrorMessage(
          `❌ Failed to launch Documentation Slayer GUI: ${err.message}`
        );
      }
      return;
    }
    if (stderr) {
      vscode.window.showWarningMessage(`⚠️ Parser stderr:\n${stderr}`);
    }
    // GUI runs independently; no further action here
  });
}

export function deactivate() {}