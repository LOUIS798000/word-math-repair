const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');

console.log('[DEBUG] 1. Script started');

// 1. Configure hardware acceleration and sandbox safely
app.disableHardwareAcceleration();
app.commandLine.appendSwitch('no-sandbox');
app.commandLine.appendSwitch('disable-gpu');

console.log('[DEBUG] 2. Flags appended');

// 2. Set safe isolated userData path in standard OS AppData
try {
  const userDataPath = path.join(app.getPath('appData'), 'word-formula-repair-tool');
  if (!fs.existsSync(userDataPath)) {
    fs.mkdirSync(userDataPath, { recursive: true });
  }
  app.setPath('userData', userDataPath);
  console.log('[DEBUG] 3. UserData set to:', userDataPath);
} catch (e) {
  console.error('Failed to set userData path:', e);
}

// 3. Ensure single instance safely
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  console.log('[INFO] Another instance is already running. Requesting focus and exiting secondary instance.');
  app.quit();
  return;
}

let mainWindow = null;

function createWindow() {
  const iconPath = path.join(__dirname, 'icon.png');
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1050,
    minHeight: 680,
    title: 'Word 公式修复工具',
    icon: iconPath,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    },
    autoHideMenuBar: true,
    show: true
  });
  console.log('[DEBUG] 7. BrowserWindow constructed');

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
  console.log('[DEBUG] 8. loadFile called');

  mainWindow.webContents.on('did-finish-load', () => {
    console.log('[DEBUG] 9. did-finish-load fired!');
    if (mainWindow) {
      mainWindow.show();
      mainWindow.focus();

      if (process.argv.includes('--capture-preview')) {
        setTimeout(async () => {
          try {
            const image = await mainWindow.webContents.capturePage();
            const previewPath = path.join(__dirname, '..', 'claude_ui_preview.png');
            fs.writeFileSync(previewPath, image.toPNG());
            console.log('[DEBUG] Screenshot captured successfully to:', previewPath);
            app.quit();
          } catch (e) {
            console.error('Failed to capture page:', e);
            app.quit();
          }
        }, 2500);
      }
    }
  });

  mainWindow.webContents.on('render-process-gone', (event, detailed) => {
    console.error('[DEBUG] Renderer process gone:', detailed);
  });

  mainWindow.on('closed', () => {
    console.log('[DEBUG] 10. Window closed');
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  console.log('[DEBUG] 5. app.whenReady fired');
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('second-instance', () => {
  if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// Helper to run python engine
function runPythonEngine(args) {
  return new Promise((resolve, reject) => {
    const engineScript = path.join(__dirname, 'engine.py');
    const pythonExe = process.platform === 'win32' ? 'python' : 'python3';
    
    const proc = spawn(pythonExe, [engineScript, ...args], {
      cwd: __dirname,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' }
    });

    let stdout = '';
    let stderr = '';

    proc.stdout.on('data', (data) => {
      stdout += data.toString('utf-8');
    });

    proc.stderr.on('data', (data) => {
      stderr += data.toString('utf-8');
    });

    proc.on('close', (code) => {
      if (code !== 0) {
        let errMsg = stderr.trim() || `Python process exited with code ${code}`;
        try {
          const parsed = JSON.parse(stdout);
          if (parsed.error) errMsg = parsed.error;
        } catch (_) {}
        reject(new Error(errMsg));
      } else {
        try {
          const raw = stdout.trim();
          const firstBrace = raw.indexOf('{');
          const lastBrace = raw.lastIndexOf('}');
          const jsonStr = (firstBrace !== -1 && lastBrace !== -1)
            ? raw.substring(firstBrace, lastBrace + 1)
            : raw;
          const result = JSON.parse(jsonStr);
          resolve(result);
        } catch (e) {
          reject(new Error(`解析引擎输出失败: ${stdout.substring(0, 300)}\n错误: ${stderr}`));
        }
      }
    });

    proc.on('error', (err) => {
      reject(new Error(`无法启动 Python 进程: ${err.message}。请确保系统已安装 Python。`));
    });
  });
}

// IPC Handlers
ipcMain.handle('dialog:openFile', async (event, allowMultiple) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择 Word 文档 (.docx)',
    filters: [{ name: 'Word 文档', extensions: ['docx'] }],
    properties: allowMultiple ? ['openFile', 'multiSelections'] : ['openFile']
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }
  return allowMultiple ? result.filePaths : result.filePaths[0];
});

ipcMain.handle('dialog:saveFile', async (event, defaultName) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '保存修复后的 Word 文档',
    defaultPath: defaultName || '已修复文档.docx',
    filters: [{ name: 'Word 文档', extensions: ['docx'] }]
  });

  if (result.canceled) return null;
  return result.filePath;
});

ipcMain.handle('engine:analyze', async (event, filePath, options) => {
  return await runPythonEngine([
    '--mode', 'analyze',
    '--input', filePath,
    '--options', JSON.stringify(options || {})
  ]);
});

ipcMain.handle('engine:repair', async (event, filePath, outputPath, options, exportPdf = false) => {
  const args = [
    '--mode', 'repair',
    '--input', filePath,
    '--output', outputPath,
    '--options', JSON.stringify(options || {})
  ];
  if (exportPdf) {
    args.push('--export-pdf');
  }
  return await runPythonEngine(args);
});

ipcMain.handle('engine:renderPages', async (event, filePath) => {
  const renderDir = path.join(__dirname, 'rendered_pages');
  return await runPythonEngine([
    '--mode', 'render_pages',
    '--input', filePath,
    '--render-dir', renderDir
  ]);
});

ipcMain.handle('app:openInWord', async (event, filePath) => {
  if (filePath && fs.existsSync(filePath)) {
    shell.openPath(filePath);
    return true;
  }
  return false;
});

ipcMain.handle('shell:openPath', async (event, targetPath) => {
  if (fs.existsSync(targetPath)) {
    if (fs.lstatSync(targetPath).isDirectory()) {
      shell.openPath(targetPath);
    } else {
      shell.showItemInFolder(targetPath);
    }
    return true;
  }
  return false;
});

ipcMain.handle('app:getDefaultSamplePath', async () => {
  const candidates = [
    path.join(__dirname, 'sample.docx'),
    path.join(__dirname, '..', 'sample.docx'),
    path.join(__dirname, '..', '模拟卷_校对排版精修版.docx')
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      return p;
    }
  }
  return null;
});
