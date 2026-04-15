const { app, BrowserWindow } = require('electron');
const path = require('path');

// URL dẫn tới thư mục chứa giao diện trên GitHub của bạn
// Vui lòng thay 'user' và 'repo/folder' bằng thông tin thật của bạn
const REMOTE_URL = "https://buiquangtrung2012-ops.github.io/HoSoDauVaoApp/taskpane.html";

function createWindow () {
  const win = new BrowserWindow({
    width: 1100,
    height: 800,
    title: "WordHoSo Desktop - Hệ Thống Tự Động",
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: true,
      contextIsolation: false,
      webSecurity: false // Cho phép load dữ liệu linh hoạt
    }
  });

  // PHƯƠNG PHÁP TỰ ĐỘNG CẬP NHẬT:
  // Thử tải từ GitHub trước để luôn có bản mới nhất
  win.loadURL(REMOTE_URL).catch(() => {
    console.log("Không có internet hoặc URL sai, nạp bản local...");
    win.loadFile(path.join(__dirname, 'public/taskpane.html'));
  });

  // Tùy chọn: Ẩn menu mặc định để nhìn chuyên nghiệp hơn
  // win.setMenu(null);
}

app.whenReady().then(() => {
  createWindow();
  
  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});
