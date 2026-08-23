const vscode = require('vscode');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
    // 1. Register Custom Editor Provider for .docx
    const provider = new DocxAgentEditorProvider(context);
    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider('docxAgent.visualEditor', provider, {
            webviewOptions: { retainContextWhenHidden: true },
            supportsMultipleEditorsPerDocument: false
        })
    );

    // 2. Register Context Menu & Command: Open in Visual Workspace
    const openCmd = vscode.commands.registerCommand('docxAgent.openWorkspace', (uri) => {
        const targetUri = uri || (vscode.window.activeTextEditor && vscode.window.activeTextEditor.document.uri);
        if (targetUri) {
            vscode.commands.executeCommand('vscode.openWith', targetUri, 'docxAgent.visualEditor');
        } else {
            vscode.window.showInformationMessage('Please select a .docx file to open with Docx-Agent.');
        }
    });
    context.subscriptions.push(openCmd);
}

class DocxAgentEditorProvider {
    constructor(context) {
        this.context = context;
    }

    async resolveCustomEditor(document, webviewPanel, _token) {
        webviewPanel.webview.options = {
            enableScripts: true,
            localResourceRoots: [this.context.extensionUri]
        };

        const filePath = document.uri.fsPath;
        const htmlPath = path.join(__dirname, '..', '..', 'src', 'docx_agent', 'interfaces', 'workspace', 'app.html');
        
        let htmlContent = '';
        if (fs.existsSync(htmlPath)) {
            htmlContent = fs.readFileSync(htmlPath, 'utf8');
        } else {
            htmlContent = `<!DOCTYPE html><html><body><h2>Docx-Agent Visual Workspace</h2><p>Loading: ${filePath}...</p></body></html>`;
        }

        webviewPanel.webview.html = htmlContent;

        // Handle bi-directional messages between VS Code webview and Python Engine
        webviewPanel.webview.onDidReceiveMessage(message => {
            switch (message.command) {
                case 'executeAgentAction':
                    vscode.window.showInformationMessage(`Docx-Agent: ${message.action}`);
                    break;
                case 'exportDocx':
                    vscode.window.showInformationMessage('Exporting DOCX...');
                    break;
            }
        });
    }
}

function deactivate() {}

module.exports = {
    activate,
    deactivate
};
