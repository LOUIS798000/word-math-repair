const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');

// 1. Configure hardware acceleration and sandbox safely
app.disableHardwareAcceleration();
app.commandLine.appendSwitch('no-sandbox');
app.commandLine.appendSwitch('disable-gpu');

// 2. Set safe isolated userData path in standard OS AppData
try {
  const userDataPath = path.join(app.getPath('appData'), 'word-formula-repair-tool');
  if (!fs.existsSync(userDataPath)) {
    fs.mkdirSync(userDataPath, { recursive: true });
  }
  app.setPath('userData', userDataPath);
} catch (e) {
  console.error('Failed to set userData path:', e);
}

// 3. Ensure single instance safely
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
  return;
}

let mainWindow = null;
let pendingOpenPath = null;

// 支持通过「打开方式」/命令行直接传入 .docx
function extractDocxPath(argv) {
  if (!Array.isArray(argv)) return null;
  for (const arg of argv) {
    if (typeof arg === 'string' && !arg.startsWith('-') && /\.docx$/i.test(arg)) {
      return path.resolve(arg);
    }
  }
  return null;
}

pendingOpenPath = extractDocxPath(process.argv.slice(1));

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

  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  mainWindow.webContents.on('did-finish-load', () => {
    if (mainWindow) {
      mainWindow.show();
      mainWindow.focus();

      if (process.argv.includes('--capture-preview')) {
        mainWindow.webContents.on('console-message', (_e, level, message) => {
          if (level >= 2) console.log('[RENDERER]', message);
        });
        setTimeout(async () => {
          try {
            const image = await mainWindow.webContents.capturePage();
            const previewPath = path.join(app.getPath('userData'), 'ui_preview.png');
            fs.writeFileSync(previewPath, image.toPNG());
            app.quit();
          } catch (e) {
            console.error('Failed to capture page:', e);
            app.quit();
          }
        }, 6000);
      }
    }
  });

  mainWindow.webContents.on('render-process-gone', (event, detailed) => {
    console.error('[DEBUG] Renderer process gone:', detailed);
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('second-instance', (event, commandLine) => {
  if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
    const file = extractDocxPath(commandLine.slice(1));
    if (file) mainWindow.webContents.send('app:openFile', file);
  }
});

app.on('open-file', (event, filePath) => {
  event.preventDefault();
  if (mainWindow) {
    mainWindow.webContents.send('app:openFile', filePath);
  } else {
    pendingOpenPath = filePath;
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// Helper to run python engine
function runPythonEngine(args) {
  return new Promise((resolve, reject) => {
    // 打包后使用随安装包附带的 engine.exe (PyInstaller 构建)，无需用户安装 Python；
    // 开发模式下直接调用系统 Python 运行 engine.py
    const bundledEngine = path.join(process.resourcesPath || '', 'engine', 'engine.exe');
    let cmd, cmdArgs, cwd;
    if (app.isPackaged && fs.existsSync(bundledEngine)) {
      cmd = bundledEngine;
      cmdArgs = args;
      cwd = path.dirname(bundledEngine);
    } else {
      cmd = process.platform === 'win32' ? 'python' : 'python3';
      cmdArgs = [path.join(__dirname, 'engine.py'), ...args];
      cwd = __dirname;
    }

    const proc = spawn(cmd, cmdArgs, {
      cwd,
      windowsHide: true,
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
      reject(new Error(app.isPackaged
        ? `无法启动内置修复引擎: ${err.message}`
        : `无法启动 Python 进程: ${err.message}。请确保系统已安装 Python。`));
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

ipcMain.handle('dialog:selectDirectory', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择批量修复结果的保存文件夹',
    properties: ['openDirectory', 'createDirectory']
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
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

function uniqueOutputPath(dir, baseName) {
  let candidate = path.join(dir, `${baseName}_修复.docx`);
  let index = 2;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${baseName}_修复_${index}.docx`);
    index += 1;
  }
  return candidate;
}

ipcMain.handle('engine:repairBatch', async (event, filePaths, outputDir, options, exportPdf = false) => {
  if (!Array.isArray(filePaths) || filePaths.length === 0) {
    throw new Error('未选择需要修复的文档');
  }
  fs.mkdirSync(outputDir, { recursive: true });

  const results = [];
  for (let i = 0; i < filePaths.length; i += 1) {
    const filePath = filePaths[i];
    const fileName = path.basename(filePath);
    event.sender.send('batch:progress', {
      index: i + 1,
      total: filePaths.length,
      fileName
    });

    const targetPath = uniqueOutputPath(outputDir, path.basename(filePath, path.extname(filePath)));
    const args = [
      '--mode', 'repair',
      '--input', filePath,
      '--output', targetPath,
      '--options', JSON.stringify(options || {})
    ];
    if (exportPdf) args.push('--export-pdf');

    try {
      const result = await runPythonEngine(args);
      results.push({
        file: filePath,
        ok: true,
        outputPath: result.savedTo || targetPath,
        totalIssues: result.totalIssues || 0
      });
    } catch (err) {
      results.push({ file: filePath, ok: false, error: err.message });
    }
  }
  return { results };
});

ipcMain.handle('engine:renderPages', async (event, filePath) => {
  // 安装目录 (app.asar) 只读，渲染图片写到用户数据目录
  const renderDir = path.join(app.getPath('userData'), 'rendered_pages');
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

ipcMain.handle('app:getLaunchFile', async () => {
  const file = pendingOpenPath;
  pendingOpenPath = null;
  return file;
});

ipcMain.handle('app:getDefaultSamplePath', async () => {
  const candidates = [
    path.join(process.resourcesPath || '', 'sample.docx'),
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
