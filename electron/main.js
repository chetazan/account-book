const { app, BrowserWindow } = require('electron')
const path = require('path')
const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: false,
      webSecurity: true
    },
    icon: path.join(__dirname, '../build/icon.ico'),
    show: false
  })

  // 개발 모드에서는 Vite 개발 서버, 프로덕션에서는 빌드된 파일
  if (isDev) {
    win.loadURL('http://localhost:5173')
    win.webContents.openDevTools()
  } else {
    win.loadFile(path.join(__dirname, '../dist/index.html'))
  }

  // 창이 준비되면 표시
  win.once('ready-to-show', () => {
    win.show()
  })

  // 개발 모드에서 Vite 서버가 준비될 때까지 대기
  if (isDev) {
    win.webContents.on('did-fail-load', () => {
      setTimeout(() => {
        win.loadURL('http://localhost:5173')
      }, 1000)
    })
  }
}

// 앱이 준비되면 창 생성
app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

// 모든 창이 닫히면 종료
app.on('window-all-closed', () => {
  app.quit()
})

// 보안: 새 창 열기 방지
app.on('web-contents-created', (event, contents) => {
  contents.on('new-window', (event, navigationUrl) => {
    event.preventDefault()
  })
})



