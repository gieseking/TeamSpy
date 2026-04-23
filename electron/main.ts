import path from 'node:path'
import { app, BrowserWindow, ipcMain } from 'electron'
import { AuthManager } from './auth'
import { getUserPhoto, loadDirectoryData } from './graph'
import type { TeamsAction } from '../src/shared/types'

let mainWindow: BrowserWindow | null = null
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
    title: 'TeamSpy',
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

app.whenReady().then(async () => {
  app.setName('TeamSpy')
  authManager = new AuthManager()

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
