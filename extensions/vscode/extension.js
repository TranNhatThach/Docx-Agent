const vscode = require('vscode');
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
    // 1. Register Default Custom Editor Provider for all .docx files
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
        
        let htmlTemplate = fs.existsSync(htmlPath) ? fs.readFileSync(htmlPath, 'utf8') : '<h2>Docx-Agent</h2>';

        // Pre-render document content using Python DocxImporter for instant display
        let renderedHtml = htmlTemplate;
        try {
            const pyScript = `import json; from pathlib import Path; from docx_agent.adapters.docx import DocxImporter; from docx_agent.canonical.model import HeadingBlock, ParagraphBlock, TableBlock; doc = DocxImporter.import_docx(r'''${filePath}'''); headings = [{'level': min(b.level, 3), 'text': b.full_text, 'id': b.id} for sec in doc.sections for b in sec.blocks if isinstance(b, HeadingBlock)]; body = ''.join([f'<h{min(b.level,3)} id=\\'{b.id}\\'>{b.full_text}</h{min(b.level,3)}>' if isinstance(b, HeadingBlock) else (f'<p id=\\'{b.id}\\'>{b.full_text}</p>' if isinstance(b, ParagraphBlock) and b.full_text.strip() else (f'<table id=\\'{b.id}\\'>{''.join(['<tr>' + ''.join([f'<td>{c.text}</td>' for c in row]) + '</tr>' for row in b.cells])}</table>' if isinstance(b, TableBlock) else '')) for sec in doc.sections for b in sec.blocks]); print(json.dumps({'title': doc.title or Path(r'''${filePath}''').stem, 'headings': headings, 'body_html': body}, ensure_ascii=False))`;
            
            const rawOutput = execSync(`python -c "${pyScript}"`, { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024 });
            const docData = JSON.parse(rawOutput);

            if (docData && docData.body_html) {
                let outlineHtml = '';
                docData.headings.forEach(h => {
                    outlineHtml += `<div class="tree-node h${h.level}" onclick="document.getElementById('${h.id}').scrollIntoView({behavior:'smooth'})">${h.text}</div>`;
                });

                renderedHtml = renderedHtml
                    .replace('<div class="outline-tree" id="outlineList">', `<div class="outline-tree" id="outlineList">${outlineHtml}`)
                    .replace('<h1>Đang tải nội dung tài liệu...</h1>', docData.body_html)
                    .replace('Bai_Tap_Oracle_HR_Schema.docx', (docData.title || path.basename(filePath)) + '.docx')
                    .replace('Đang tải...', `${docData.headings.length} mục`);
            }
        } catch (err) {
            console.error('Docx-Agent pre-render error:', err);
        }

        // Set Webview HTML inside Antigravity Tab
        webviewPanel.webview.html = renderedHtml;

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
