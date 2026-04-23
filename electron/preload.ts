import { contextBridge, ipcRenderer } from 'electron'
import type {
  AppSettings,
  AuthState,
  DirectoryPayload,
  TeamSpyDesktopApi,
  TeamsAction,
} from '../src/shared/types'

const api: TeamSpyDesktopApi = {
  auth: {
    getState: () => ipcRenderer.invoke('auth:get-state') as Promise<AuthState>,
    login: () => ipcRenderer.invoke('auth:login') as Promise<AuthState>,
    logout: () => ipcRenderer.invoke('auth:logout') as Promise<AuthState>,
    saveSettings: (settings: AppSettings) =>
      ipcRenderer.invoke('auth:save-settings', settings) as Promise<AuthState>,
  },
  directory: {
    load: () =>
      ipcRenderer.invoke('directory:load') as Promise<DirectoryPayload>,
    getPhoto: (userId: string) =>
      ipcRenderer.invoke('directory:get-photo', userId) as Promise<string | null>,
  },
  teams: {
    openAction: (action: TeamsAction, email: string) =>
      ipcRenderer.invoke('teams:open-action', action, email) as Promise<boolean>,
  },
}

contextBridge.exposeInMainWorld('teamspy', api)
