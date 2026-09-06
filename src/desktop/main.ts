import { app, BrowserWindow, dialog, ipcMain } from 'electron';
import { readFile } from 'node:fs/promises';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { Workspace, type TextEdit } from '../engine/workspace';
import { saveAs } from '../io/save';
import { DocxPackage } from '../package/docx';
import { readDocument } from '../engine/document';

const workspace = new Workspace();
let window: BrowserWindow;
const uiFile = join(__dirname, '../ui/index.html');
const uiUrl = pathToFileURL(uiFile).href;
const snapshot = () => ({
  documents: workspace.documents(),
  canUndo: workspace.canUndo,
  canRedo: workspace.canRedo,
});
const string = (value: unknown): string => {
  if (typeof value !== 'string' || value.length > 1_000_000)
    throw new Error('Expected a bounded string');
  return value;
};

function handle(name: string, handler: (...args: any[]) => unknown) {
  ipcMain.handle(`ritr:${name}`, (event, ...args) => {
    if (
      event.sender !== window.webContents ||
      event.senderFrame !== window.webContents.mainFrame ||
      event.senderFrame.url !== uiUrl
    )
      throw new Error('Untrusted IPC sender');
    return handler(...args);
  });
}

async function createWindow() {
  window = new BrowserWindow({
    width: 1440,
    height: 940,
    minWidth: 1050,
    minHeight: 700,
    title: 'ritr — DOCX workbench',
    backgroundColor: '#f4f1e9',
    show: process.env.RITR_SMOKE !== '1',
    webPreferences: {
      preload: join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      offscreen: process.env.RITR_SMOKE === '1',
      backgroundThrottling: false,
    },
  });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event) => event.preventDefault());
  window.webContents.session.setPermissionRequestHandler((_contents, _permission, callback) =>
    callback(false),
  );
  window.on('close', (event) => {
    if (process.env.RITR_SMOKE !== '1' && workspace.documents().some((d) => d.dirty)) {
      const choice = dialog.showMessageBoxSync(window, {
        type: 'question',
        buttons: ['Keep editing', 'Discard and close'],
        defaultId: 0,
        cancelId: 0,
        message: 'Close with unsaved changes?',
        detail: 'Save As each changed document before closing to keep your edits.',
      });
      if (choice === 0) event.preventDefault();
    }
  });
  await window.loadFile(uiFile);
}

app
  .whenReady()
  .then(async () => {
    handle('snapshot', snapshot);
    handle('open', async () => {
      const result = await dialog.showOpenDialog(window, {
        properties: ['openFile', 'multiSelections'],
        filters: [{ name: 'Word documents', extensions: ['docx'] }],
      });
      const files = await Promise.all(
        result.filePaths.map(async (file) => ({ file, bytes: await readFile(file) })),
      );
      for (const { bytes } of files) readDocument(DocxPackage.open(bytes));
      for (const { file, bytes } of files) workspace.open(basename(file), bytes);
      return snapshot();
    });
    handle('previewEdit', (value: TextEdit) => {
      if (
        !value ||
        typeof value !== 'object' ||
        !Number.isInteger(value.from) ||
        !Number.isInteger(value.to)
      )
        throw new Error('Invalid text edit');
      return workspace.preview('Edit text', [
        {
          documentId: string(value.documentId),
          spanId: string(value.spanId),
          from: value.from,
          to: value.to,
          text: string(value.text),
        },
      ]);
    });
    handle('previewReplace', (query, replacement, sensitive) =>
      workspace.previewReplace(string(query), string(replacement), sensitive !== false),
    );
    handle('search', (query, sensitive) => workspace.search(string(query), sensitive !== false));
    handle('commit', (id) => {
      workspace.commit(string(id));
      return snapshot();
    });
    handle('undo', () => {
      workspace.undo();
      return snapshot();
    });
    handle('redo', () => {
      workspace.redo();
      return snapshot();
    });
    handle('report', (id) => workspace.report(string(id)));
    handle('save', async (value) => {
      const id = string(value);
      const pkg = workspace.package(id);
      const doc = workspace.document(id);
      const result = await dialog.showSaveDialog(window, {
        defaultPath: doc.name.replace(/\.docx$/i, '') + '-edited.docx',
        filters: [{ name: 'Word documents', extensions: ['docx'] }],
      });
      if (!result.canceled && result.filePath) {
        await saveAs(result.filePath, pkg);
        workspace.markSaved(id, pkg);
      }
      return {
        path: result.canceled ? undefined : result.filePath,
        snapshot: snapshot(),
        report: workspace.report(id),
      };
    });
    // Explicit startup paths also make end-to-end testing possible without automating OS dialogs.
    const files = process.argv
      .slice(app.isPackaged ? 1 : 2)
      .filter((arg) => arg.toLowerCase().endsWith('.docx'));
    for (const file of files) workspace.open(basename(file), await readFile(resolve(file)));
    await createWindow();
    app.on('activate', () => {
      if (!BrowserWindow.getAllWindows().length) void createWindow();
    });
  })
  .catch((error) => {
    console.error(error);
    app.exit(1);
  });
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
