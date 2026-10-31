// शोध Mac ॲप: स्वतःची खिडकी, आत तोच स्थानिक सर्व्हर. ब्राउझर लागत नाही.
const { app, BrowserWindow, dialog, shell, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');

process.env.SHODH_DATA = app.getPath('userData');
const { start } = require('./server');

let win;

// Ollama बसवलेले असेल पण बंद असेल तर आपोआप सुरू करा
async function ensureOllama() {
  try {
    await fetch('http://127.0.0.1:11434/api/tags', { signal: AbortSignal.timeout(1000) });
  } catch {
    if (fs.existsSync('/Applications/Ollama.app')) execFile('open', ['-g', '-a', 'Ollama']);
  }
}

// निकालाचे पान लपलेल्या खिडकीत उघडून खरी PDF फाइल बनवणे
async function savePdf(url) {
  const w = new BrowserWindow({ show: false });
  try {
    await w.loadURL(url + (url.includes('?') ? '&' : '?') + 'app=1');
    const title = w.getTitle().replace(/[\\/:*?"<>|]+/g, ' ').slice(0, 60) || 'shodh';
    const data = await w.webContents.printToPDF({ pageSize: 'A4', printBackground: true, preferCSSPageSize: true, margins: { marginType: 'default' } });
    const { canceled, filePath } = await dialog.showSaveDialog(win, {
      defaultPath: path.join(app.getPath('downloads'), title + '.pdf'),
      filters: [{ name: 'PDF', extensions: ['pdf'] }]
    });
    if (!canceled && filePath) {
      fs.writeFileSync(filePath, data);
      shell.openPath(filePath);
    }
  } finally {
    w.destroy();
  }
}

function createWindow(port) {
  win = new BrowserWindow({
    width: 1100,
    height: 820,
    minWidth: 600,
    title: 'शोध',
    backgroundColor: '#f6f7f9',
    webPreferences: { preload: path.join(__dirname, 'preload.js') }
  });
  win.loadURL(`http://localhost:${port}/`);

  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/\/api\/download\/[0-9a-f]+\.pdf/.test(url)) savePdf(url).catch(e => dialog.showErrorBox('PDF', e.message));
    else if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  // Word / Excel: कुठे जतन करायचे ते विचारून, जतन झाल्यावर उघडणे
  win.webContents.session.on('will-download', (_e, item) => {
    item.setSaveDialogOptions({ defaultPath: path.join(app.getPath('downloads'), item.getFilename()) });
    item.once('done', (_ev, state) => { if (state === 'completed') shell.openPath(item.getSavePath()); });
  });
}

ipcMain.handle('pick-folder', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openDirectory'], message: 'माहितीचा फोल्डर निवडा' });
  return r.canceled ? null : r.filePaths[0];
});

if (!app.requestSingleInstanceLock()) app.quit();
app.on('second-instance', () => { if (win) { win.show(); win.focus(); } });

app.whenReady().then(async () => {
  const icon = path.join(__dirname, 'build', 'icon.png');
  if (app.dock && fs.existsSync(icon)) app.dock.setIcon(icon);
  ensureOllama();
  let port;
  try { port = await start(3939); } catch { port = await start(0); }
  createWindow(port);
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(port); });
});

app.on('window-all-closed', () => app.quit());
