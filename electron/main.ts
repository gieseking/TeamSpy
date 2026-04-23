import path from 'node:path'
import { app, BrowserWindow, ipcMain, shell } from 'electron'
import { AuthManager } from './auth'
import { getUserPhoto, loadDirectoryData } from './graph'
import type { TeamsAction } from '../src/shared/types'

let mainWindow: BrowserWindow | null = null
let authWindow: BrowserWindow | null = null
let authManager: AuthManager

function getRendererEntry() {
  const devServerUrl = process.env.VITE_DEV_SERVER_URL

  if (devServerUrl) {
    return {
      mode: 'url' as const,
      value: devServerUrl,
    }
  }

  return {
    mode: 'file' as const,
    value: path.join(app.getAppPath(), 'dist', 'index.html'),
  }
}

function buildTeamsUrl(action: TeamsAction, email: string) {
  const encoded = encodeURIComponent(email)

  if (action === 'call') {
    return `https://teams.microsoft.com/l/call/0/0?users=${encoded}`
  }

  return `https://teams.microsoft.com/l/chat/0/0?users=${encoded}`
}

async function createWindow() {
  const preloadPath = path.join(__dirname, 'preload.cjs')
  const isMac = process.platform === 'darwin'

  mainWindow = new BrowserWindow({
    width: 1560,
    height: 980,
    minWidth: 1180,
    minHeight: 720,
    title: 'FriendlyFaces',
    ...(isMac ? { titleBarStyle: 'hiddenInset' as const } : {}),
    autoHideMenuBar: !isMac,
    backgroundColor: '#e9eef6',
    webPreferences: {
      preload: preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  })

  const entry = getRendererEntry()

  if (entry.mode === 'url') {
    await mainWindow.loadURL(entry.value)
    mainWindow.webContents.openDevTools({ mode: 'detach' })
  } else {
    await mainWindow.loadFile(entry.value)
  }
}

async function openAuthWindow(url: string) {
  if (authWindow && !authWindow.isDestroyed()) {
    await authWindow.loadURL(url)
    authWindow.focus()
    return
  }

  authWindow = new BrowserWindow({
    width: 540,
    height: 760,
    minWidth: 460,
    minHeight: 620,
    title: 'Connect Microsoft 365',
    parent: mainWindow ?? undefined,
    modal: Boolean(mainWindow),
    autoHideMenuBar: true,
    backgroundColor: '#f5f8fc',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })

  authWindow.on('closed', () => {
    authWindow = null
  })

  authWindow.webContents.setWindowOpenHandler(({ url: targetUrl }) => {
    void authWindow?.loadURL(targetUrl)
    return { action: 'deny' }
  })

  authWindow.webContents.on('will-navigate', (event, targetUrl) => {
    if (targetUrl.startsWith('https://teams.microsoft.com/')) {
      event.preventDefault()
      void shell.openExternal(targetUrl)
    }
  })

  authWindow.webContents.on(
    'did-fail-load',
    (_event, errorCode, _errorDescription, validatedUrl) => {
      if (errorCode === -3 || validatedUrl.startsWith('http://localhost')) {
        return
      }
    },
  )

  await authWindow.loadURL(url)
}

function closeAuthWindow() {
  if (!authWindow || authWindow.isDestroyed()) {
    authWindow = null
    return
  }

  authWindow.close()
  authWindow = null
}

app.whenReady().then(async () => {
  app.setName('FriendlyFaces')
  authManager = new AuthManager({
    open: openAuthWindow,
    close: closeAuthWindow,
  })

  ipcMain.handle('auth:get-state', () => authManager.getAuthState())
  ipcMain.handle('auth:login', () => authManager.login())
  ipcMain.handle('auth:logout', () => authManager.logout())
  ipcMain.handle('directory:load', () => loadDirectoryData(authManager))
  ipcMain.handle('directory:get-photo', (_event, userId: string) =>
    getUserPhoto(authManager, userId),
  )
  ipcMain.handle(
    'teams:open-action',
    async (_event, action: TeamsAction, email: string) => {
      if (!email) {
        return false
      }

      await authManager.openExternal(buildTeamsUrl(action, email))
      return true
    },
  )

  await createWindow()

  app.on('activate', async () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      await createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
