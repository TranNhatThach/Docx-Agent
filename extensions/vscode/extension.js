const vscode = require('vscode');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
    // 1. Register Custom Editor Provider for .docx inside VS Code / Antigravity
    const provider = new DocxAgentEditorProvider(context);
    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider('docxAgent.visualEditor', provider, {
            webviewOptions: { retainContextWhenHidden: true },
            supportsMultipleEditorsPerDocument: false
        })
    );

    // 2. Register Command: Open in Visual Workspace
    const openCmd = vscode.commands.registerCommand('docxAgent.openWorkspace', (uri) => {
        const targetUri = uri || (vscode.window.activeTextEditor && vscode.window.activeTextEditor.document.uri);
        if (targetUri) {
            vscode.commands.executeCommand('vscode.openWith', targetUri, 'docxAgent.visualEditor');
        } else {
            vscode.window.showInformationMessage('Vui lòng chọn một file .docx để mở trực tiếp trong Antigravity.');
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
        let htmlPath = path.join(__dirname, 'app.html');
        if (!fs.existsSync(htmlPath)) {
            htmlPath = path.join(__dirname, '..', '..', 'src', 'docx_agent', 'interfaces', 'workspace', 'app.html');
        }
        
        let htmlContent = fs.existsSync(htmlPath) ? fs.readFileSync(htmlPath, 'utf8') : '<h2>Docx-Agent</h2>';

        // Set Webview HTML inside Antigravity Tab
        webviewPanel.webview.html = htmlContent;

        // Fetch document outline & content from docx-agent CLI
        exec(`docx-agent inspect "${filePath}" --json`, (err, stdout) => {
            if (!err && stdout) {
                try {
                    const data = JSON.parse(stdout);
                    webviewPanel.webview.postMessage({ command: 'loadDocument', data: data });
                } catch(e) {}
            }
        });

        // Handle bi-directional messages from Webview to Antigravity
        webviewPanel.webview.onDidReceiveMessage(message => {
            switch (message.command) {
                case 'executeAgentAction':
                    vscode.window.showInformationMessage(`Docx-Agent: ${message.action}`);
                    break;
                case 'exportDocx':
                    vscode.window.showInformationMessage('Đang lưu và xuất file DOCX...');
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
