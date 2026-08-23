const vscode = require('vscode');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
    // 1. Register Default Custom Editor Provider for .docx files
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

    // 3. Register Command: Get Active Selection Context
    const getSelCmd = vscode.commands.registerCommand('docxAgent.getSelectionContext', () => {
        return provider.getActiveSelection();
    });
    context.subscriptions.push(getSelCmd);
}

class DocxAgentEditorProvider {
    constructor(context) {
        this.context = context;
        this.activeSelection = null;
    }

    getActiveSelection() {
        return this.activeSelection;
    }

    async openCustomDocument(uri, _openContext, _token) {
        return {
            uri: uri,
            dispose: () => {}
        };
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

        let htmlContent = fs.existsSync(htmlPath) ? fs.readFileSync(htmlPath, 'utf8') : '<h2>Docx-Agent V2.1</h2>';

        // Inject bundled offline libraries
        const jszipPath = path.join(__dirname, 'jszip.min.js');
        const docxPrevPath = path.join(__dirname, 'docx-preview.min.js');
        if (fs.existsSync(jszipPath)) {
            const jszipCode = fs.readFileSync(jszipPath, 'utf8');
            htmlContent = htmlContent.replace('<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>', `<script>${jszipCode}</script>`);
        }
        if (fs.existsSync(docxPrevPath)) {
            const docxPrevCode = fs.readFileSync(docxPrevPath, 'utf8');
            htmlContent = htmlContent.replace('<script src="https://cdn.jsdelivr.net/npm/docx-preview@0.4.0/dist/docx-preview.min.js"></script>', `<script>${docxPrevCode}</script>`);
        }

        webviewPanel.webview.html = htmlContent;

        // Find Python binary and workspace root dynamically
        const pythonBin = this.getPythonExecutable(filePath);
        const workspaceSrc = this.findDocxAgentSrc(filePath);

        // Handle Bi-directional message passing
        webviewPanel.webview.onDidReceiveMessage(async (message) => {
            if (!message || !message.type) return;

            switch (message.type) {
                case 'WEBVIEW_READY': {
                    webviewPanel.webview.postMessage({
                        type: 'DOCUMENT_LOADING_STATE',
                        stage: 'LOADING',
                        message: 'Đang tải và phân tích tệp OpenXML...'
                    });

                    this.loadDocumentAsync(pythonBin, workspaceSrc, filePath, webviewPanel);
                    break;
                }

                case 'SELECTION_CHANGED': {
                    this.activeSelection = message.payload;
                    this.persistActiveSelection(message.payload, filePath);
                    break;
                }

                case 'REQUEST_AGENT_ACTION': {
                    const action = message.payload.action;
                    const selText = message.payload.context ? message.payload.context.selected_text : '';
                    vscode.window.showInformationMessage(`Antigravity Agent: [${action}] "${selText.substring(0, 40)}..."`);
                    break;
                }

                case 'SAVE_DOCUMENT': {
                    this.saveDocumentAsync(pythonBin, workspaceSrc, filePath, message.payload, webviewPanel);
                    break;
                }

                case 'APPLY_TRANSACTION': {
                    vscode.window.showInformationMessage(`Đã áp dụng giao dịch: ${message.payload.transaction_id}`);
                    break;
                }

                case 'ACCEPT_CHANGE': {
                    const changeId = message.payload.change_id;
                    vscode.window.setStatusBarMessage(`✓ Đã chấp nhận thay đổi ${changeId}`, 3000);
                    break;
                }

                case 'REJECT_CHANGE': {
                    const changeId = message.payload.change_id;
                    vscode.window.setStatusBarMessage(`✕ Đã từ chối thay đổi ${changeId}`, 3000);
                    break;
                }

                case 'ACCEPT_ALL_CHANGES': {
                    vscode.window.showInformationMessage(`Đã chấp nhận toàn bộ thay đổi trong phiên AI.`);
                    break;
                }

                case 'REJECT_ALL_CHANGES': {
                    vscode.window.showInformationMessage(`Đã từ chối toàn bộ thay đổi trong phiên AI.`);
                    break;
                }

                case 'COMMIT_SESSION': {
                    vscode.window.showInformationMessage(`Đang commit phiên thay đổi vào phiên bản tài liệu mới...`);
                    break;
                }

                case 'GET_VERSION_HISTORY': {
                    this.fetchVersionHistoryAsync(pythonBin, workspaceSrc, filePath, webviewPanel);
                    break;
                }

                case 'RESTORE_VERSION': {
                    const vNum = message.payload.version;
                    vscode.window.showInformationMessage(`Đang khôi phục tài liệu về Phiên bản ${vNum}...`);
                    break;
                }
            }
        });
    }

    findDocxAgentSrc(filePath) {
        // 1. Walk up from filePath
        if (filePath) {
            let cur = path.dirname(filePath);
            for (let i = 0; i < 6; i++) {
                const c1 = path.join(cur, 'Docx-Agent', 'src');
                if (fs.existsSync(path.join(c1, 'docx_agent'))) return c1;
                const c2 = path.join(cur, 'src');
                if (fs.existsSync(path.join(c2, 'docx_agent'))) return c2;
                const parent = path.dirname(cur);
                if (parent === cur) break;
                cur = parent;
            }
        }

        // 2. Search in workspace folders
        if (vscode.workspace.workspaceFolders) {
            for (const wf of vscode.workspace.workspaceFolders) {
                const c1 = path.join(wf.uri.fsPath, 'Docx-Agent', 'src');
                if (fs.existsSync(path.join(c1, 'docx_agent'))) return c1;
                const c2 = path.join(wf.uri.fsPath, 'src');
                if (fs.existsSync(path.join(c2, 'docx_agent'))) return c2;
            }
        }

        // 3. Fallback to extension directory parent
        const devSrc = path.join(__dirname, '..', '..', 'src');
        if (fs.existsSync(path.join(devSrc, 'docx_agent'))) return devSrc;

        return '';
    }

    loadDocumentAsync(pythonBin, workspaceSrc, filePath, webviewPanel) {
        const startTime = Date.now();
        const args = ['-m', 'docx_agent.interfaces.cli.main', 'workspace-load', filePath, '--json'];

        const env = Object.assign({}, process.env, {
            PYTHONPATH: workspaceSrc ? `${workspaceSrc}${path.delimiter}${process.env.PYTHONPATH || ''}` : process.env.PYTHONPATH,
            PYTHONIOENCODING: 'utf-8'
        });

        const proc = spawn(pythonBin, args, { env });
        let stdout = '';
        let stderr = '';

        proc.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
        proc.stderr.on('data', (d) => { stderr += d.toString('utf8'); });

        proc.on('close', (code) => {
            const elapsed = Date.now() - startTime;
            if (code === 0 && stdout.trim()) {
                try {
                    const docData = JSON.parse(stdout);
                    let docxBase64 = null;
                    try {
                        if (fs.existsSync(filePath)) {
                            docxBase64 = fs.readFileSync(filePath).toString('base64');
                        }
                    } catch (_) {}

                    webviewPanel.webview.postMessage({
                        type: 'DOCUMENT_LOADED',
                        payload: docData,
                        docxBase64: docxBase64
                    });
                } catch (e) {
                    webviewPanel.webview.postMessage({
                        type: 'DOCUMENT_LOAD_ERROR',
                        payload: {
                            stage: 'PARSING',
                            error_type: 'JSONParseError',
                            message: 'Không thể phân tích dữ liệu JSON trả về từ bộ xử lý tài liệu.',
                            diagnostics: e.message,
                            stderr: stdout.substring(0, 500)
                        }
                    });
                }
            } else {
                webviewPanel.webview.postMessage({
                    type: 'DOCUMENT_LOAD_ERROR',
                    payload: {
                        stage: 'LOADING',
                        error_type: 'PythonBridgeExecutionError',
                        message: `Bộ xử lý tài liệu trả về mã lỗi: ${code}`,
                        diagnostics: stderr || stdout || 'Không có thông báo lỗi từ tiến trình Python.',
                        stderr: stderr,
                        command: `${pythonBin} ${args.join(' ')}`,
                        elapsed_ms: elapsed,
                        file_path: filePath
                    }
                });
            }
        });

        proc.on('error', (err) => {
            webviewPanel.webview.postMessage({
                type: 'DOCUMENT_LOAD_ERROR',
                payload: {
                    stage: 'PYTHON_BRIDGE_CONNECTING',
                    error_type: 'ProcessSpawnError',
                    message: `Không thể khởi chạy Python: ${err.message}`,
                    diagnostics: err.stack,
                    command: pythonBin
                }
            });
        });
    }

    saveDocumentAsync(pythonBin, workspaceSrc, filePath, payload, webviewPanel) {
        let tmpDir = path.join(path.dirname(filePath), '.docx_agent_workspace');
        if (!fs.existsSync(tmpDir)) {
            try {
                fs.mkdirSync(tmpDir, { recursive: true });
            } catch (_) {
                tmpDir = require('os').tmpdir();
            }
        }
        const tmpPayloadFile = path.join(tmpDir, `save_payload_${Date.now()}.json`);
        fs.writeFileSync(tmpPayloadFile, JSON.stringify(payload, null, 2), 'utf8');

        const args = ['-m', 'docx_agent.interfaces.cli.main', 'workspace-save', filePath, tmpPayloadFile, '--output', filePath, '--json'];

        const env = Object.assign({}, process.env, {
            PYTHONPATH: workspaceSrc ? `${workspaceSrc}${path.delimiter}${process.env.PYTHONPATH || ''}` : process.env.PYTHONPATH,
            PYTHONIOENCODING: 'utf-8'
        });

        const proc = spawn(pythonBin, args, { env });
        let stdout = '';
        let stderr = '';

        proc.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
        proc.stderr.on('data', (d) => { stderr += d.toString('utf8'); });

        proc.on('close', (code) => {
            try { fs.unlinkSync(tmpPayloadFile); } catch (_) {}

            if (code === 0 && stdout.trim()) {
                try {
                    const resData = JSON.parse(stdout);
                    webviewPanel.webview.postMessage({
                        type: 'SAVE_SUCCESS',
                        payload: resData
                    });
                    vscode.window.setStatusBarMessage('Docx-Agent: Đã lưu tài liệu thành công', 3000);
                } catch (e) {
                    webviewPanel.webview.postMessage({
                        type: 'SAVE_ERROR',
                        payload: { message: 'Lỗi phân tích phản hồi lưu tài liệu.' }
                    });
                }
            } else {
                webviewPanel.webview.postMessage({
                    type: 'SAVE_ERROR',
                    payload: { message: stderr || 'Lỗi khi ghi file DOCX.' }
                });
            }
        });
    }

    fetchVersionHistoryAsync(pythonBin, workspaceSrc, filePath, webviewPanel) {
        const args = ['-m', 'docx_agent.interfaces.cli.main', 'version-history', filePath, '--json'];
        const env = Object.assign({}, process.env, {
            PYTHONPATH: workspaceSrc ? `${workspaceSrc}${path.delimiter}${process.env.PYTHONPATH || ''}` : process.env.PYTHONPATH,
            PYTHONIOENCODING: 'utf-8'
        });

        const proc = spawn(pythonBin, args, { env });
        let stdout = '';
        let stderr = '';

        proc.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
        proc.stderr.on('data', (d) => { stderr += d.toString('utf8'); });

        proc.on('close', (code) => {
            if (code === 0 && stdout.trim()) {
                try {
                    const resData = JSON.parse(stdout);
                    webviewPanel.webview.postMessage({
                        type: 'VERSION_HISTORY_RECEIVED',
                        payload: resData.versions || []
                    });
                } catch (_) {}
            }
        });
    }

    persistActiveSelection(selectionData, filePath) {
        try {
            let targetDir = null;
            if (vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0) {
                targetDir = path.join(vscode.workspace.workspaceFolders[0].uri.fsPath, 'Docx-Agent', '.docx_agent_workspace');
                if (!fs.existsSync(path.dirname(targetDir))) {
                    targetDir = path.join(vscode.workspace.workspaceFolders[0].uri.fsPath, '.docx_agent_workspace');
                }
            } else if (filePath) {
                targetDir = path.join(path.dirname(filePath), '.docx_agent_workspace');
            }
            if (targetDir) {
                if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });
                fs.writeFileSync(path.join(targetDir, 'active_selection.json'), JSON.stringify(selectionData, null, 2), 'utf8');
            }
        } catch (_) {}
    }

    getPythonExecutable(filePath) {
        const config = vscode.workspace.getConfiguration('python');
        const defaultInterpreter = config.get('defaultInterpreterPath');
        if (defaultInterpreter && fs.existsSync(defaultInterpreter)) {
            return defaultInterpreter;
        }

        const searchRoots = [];
        if (filePath) {
            let cur = path.dirname(filePath);
            for (let i = 0; i < 4; i++) {
                searchRoots.push(cur);
                searchRoots.push(path.join(cur, 'Docx-Agent'));
                const parent = path.dirname(cur);
                if (parent === cur) break;
                cur = parent;
            }
        }
        if (vscode.workspace.workspaceFolders) {
            for (const wf of vscode.workspace.workspaceFolders) {
                searchRoots.push(wf.uri.fsPath);
                searchRoots.push(path.join(wf.uri.fsPath, 'Docx-Agent'));
            }
        }

        for (const root of searchRoots) {
            const venvs = [
                path.join(root, '.venv', 'Scripts', 'python.exe'),
                path.join(root, 'venv', 'Scripts', 'python.exe'),
                path.join(root, '.venv', 'bin', 'python'),
                path.join(root, 'venv', 'bin', 'python')
            ];
            for (const v of venvs) {
                if (fs.existsSync(v)) return v;
            }
        }

        return 'python';
    }
}

function deactivate() {}

module.exports = {
    activate,
    deactivate
};
