const { app, BrowserWindow, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');

// ═══════════════════════════════════════════════
// 自动更新配置（GitHub Releases）
// ═══════════════════════════════════════════════
// 每次用 npm run release 构建并发布到 GitHub Releases 后，
// 用户打开软件即可自动检查更新

autoUpdater.setFeedURL({
  provider: 'github',
  owner: 'liusite66-dev',
  repo: '-traffic-compensation-calculator',
});
autoUpdater.autoDownload = false;

function checkForUpdates() {
  autoUpdater.checkForUpdates().catch(() => {
    // 无网络连接时静默忽略
  });
}

autoUpdater.on('update-available', (info) => {
  dialog.showMessageBox({
    type: 'info',
    title: '发现新版本',
    message: `有新版本 v${info.version} 可用，是否下载更新？`,
    buttons: ['下载更新', '暂不更新'],
    defaultId: 0,
    cancelId: 1,
  }).then(({ response }) => {
    if (response === 0) {
      autoUpdater.downloadUpdate();
    }
  });
});

autoUpdater.on('update-not-available', () => {});

autoUpdater.on('download-progress', (progress) => {
  // 可在此添加进度显示
});

autoUpdater.on('update-downloaded', () => {
  dialog.showMessageBox({
    type: 'info',
    title: '更新已下载',
    message: '更新已下载完成，是否立即重启以安装更新？',
    buttons: ['立即重启', '稍后重启'],
    defaultId: 0,
    cancelId: 1,
  }).then(({ response }) => {
    if (response === 0) {
      autoUpdater.quitAndInstall();
    }
  });
});

autoUpdater.on('error', () => {
  // 静默忽略更新错误
});

// ═══════════════════════════════════════════════
// 窗口创建
// ═══════════════════════════════════════════════

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 1000,
    minHeight: 650,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  win.loadFile('index.html');
  win.once('ready-to-show', () => {
    win.show();
    // 启动3秒后检查更新
    setTimeout(checkForUpdates, 3000);
  });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
