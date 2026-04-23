import { promises as fs } from 'node:fs'
import path from 'node:path'
import { app, shell } from 'electron'
import dotenv from 'dotenv'
import {
  InteractionRequiredAuthError,
  LogLevel,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type TokenCacheContext,
} from '@azure/msal-node'
import type { AppSettings, AuthAccount, AuthState } from '../src/shared/types'

const REQUIRED_SETTINGS: Array<keyof AppSettings> = ['clientId', 'tenantId']
const REDIRECT_URI = 'http://localhost'

export const GRAPH_SCOPES = [
  'User.Read',
  'User.Read.All',
  'Presence.Read.All',
  'ProfilePhoto.Read.All',
  'MailboxSettings.Read',
  'AuditLog.Read.All',
] as const

class FileCachePlugin {
  constructor(private readonly cachePath: string) {}

  async beforeCacheAccess(cacheContext: TokenCacheContext) {
    try {
      const cache = await fs.readFile(this.cachePath, 'utf8')
      cacheContext.tokenCache.deserialize(cache)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') {
        throw error
      }
    }
  }

  async afterCacheAccess(cacheContext: TokenCacheContext) {
    if (!cacheContext.cacheHasChanged) {
      return
    }

    await fs.mkdir(path.dirname(this.cachePath), { recursive: true })
    await fs.writeFile(this.cachePath, cacheContext.tokenCache.serialize(), 'utf8')
  }
}

function mapAccount(account: AccountInfo | null): AuthAccount | null {
  if (!account) {
    return null
  }

  return {
    homeAccountId: account.homeAccountId,
    name: account.name ?? account.username,
    username: account.username,
    tenantId: account.tenantId,
  }
}

async function readJsonFile<T>(filePath: string, fallback: T) {
  try {
    const raw = await fs.readFile(filePath, 'utf8')
    return JSON.parse(raw) as T
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
      return fallback
    }

    throw error
  }
}

function sanitizeSettings(input: Partial<AppSettings> | null | undefined): AppSettings {
  return {
    clientId: input?.clientId?.trim() ?? '',
    tenantId: input?.tenantId?.trim() ?? '',
  }
}

function missingSettings(settings: AppSettings) {
  return REQUIRED_SETTINGS.filter((key) => !settings[key])
}

function loadDesktopEnv() {
  const candidates = new Set([
    path.join(process.cwd(), '.env.local'),
    path.join(process.cwd(), '.env'),
    path.join(app.getPath('userData'), '.env'),
  ])

  for (const candidate of candidates) {
    dotenv.config({
      path: candidate,
      override: false,
    })
  }
}

export class AuthManager {
  private clientApplication: PublicClientApplication | null = null
  private settings: AppSettings = { clientId: '', tenantId: '' }
  private readonly settingsPath = path.join(app.getPath('userData'), 'settings.json')
  private readonly cachePath = path.join(app.getPath('userData'), 'msal-cache.json')

  async initialize() {
    loadDesktopEnv()

    const storedSettings = await readJsonFile<Partial<AppSettings>>(this.settingsPath, {})
    const envSettings = sanitizeSettings({
      clientId: process.env.TEAMSPY_CLIENT_ID,
      tenantId: process.env.TEAMSPY_TENANT_ID,
    })

    this.settings = sanitizeSettings({
      ...storedSettings,
      clientId: storedSettings.clientId || envSettings.clientId,
      tenantId: storedSettings.tenantId || envSettings.tenantId,
    })

    await this.configureClientApplication()
  }

  getAuthState = async (): Promise<AuthState> => {
    const account = await this.getAccount()

    return {
      configured: missingSettings(this.settings).length === 0,
      signedIn: account !== null,
      account: mapAccount(account),
      settings: this.settings,
      missingSettings: missingSettings(this.settings),
    }
  }

  saveSettings = async (nextSettings: AppSettings): Promise<AuthState> => {
    this.settings = sanitizeSettings(nextSettings)

    await fs.mkdir(path.dirname(this.settingsPath), { recursive: true })
    await fs.writeFile(
      this.settingsPath,
      JSON.stringify(this.settings, null, 2),
      'utf8',
    )

    await fs.rm(this.cachePath, { force: true })
    await this.configureClientApplication()

    return this.getAuthState()
  }

  login = async (): Promise<AuthState> => {
    this.ensureConfigured()

    await this.acquireTokenInteractive([...GRAPH_SCOPES])
    return this.getAuthState()
  }

  logout = async (): Promise<AuthState> => {
    const account = await this.getAccount()

    if (account && this.clientApplication) {
      await this.clientApplication.getTokenCache().removeAccount(account)
    }

    return this.getAuthState()
  }

  getAccessToken = async (scopes: string[] = [...GRAPH_SCOPES]) => {
    this.ensureConfigured()

    const account = await this.getAccount()

    if (!account) {
      const response = await this.acquireTokenInteractive(scopes)
      return response.accessToken
    }

    try {
      const response = await this.clientApplication!.acquireTokenSilent({
        account,
        scopes,
      })

      return response.accessToken
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await this.acquireTokenInteractive(scopes)
        return response.accessToken
      }

      throw error
    }
  }

  openExternal = async (url: string) => {
    await shell.openExternal(url)
  }

  private async configureClientApplication() {
    if (missingSettings(this.settings).length > 0) {
      this.clientApplication = null
      return
    }

    const config: Configuration = {
      auth: {
        clientId: this.settings.clientId,
        authority: `https://login.microsoftonline.com/${this.settings.tenantId}`,
      },
      cache: {
        cachePlugin: new FileCachePlugin(this.cachePath),
      },
      system: {
        loggerOptions: {
          logLevel: LogLevel.Warning,
          piiLoggingEnabled: false,
          loggerCallback(_level, message) {
            console.log(`[msal] ${message}`)
          },
        },
      },
    }

    this.clientApplication = new PublicClientApplication(config)
  }

  private ensureConfigured() {
    if (!this.clientApplication) {
      throw new Error(
        `Missing required settings: ${missingSettings(this.settings).join(', ')}`,
      )
    }
  }

  private getAccount = async () => {
    if (!this.clientApplication) {
      return null
    }

    const accounts = await this.clientApplication.getTokenCache().getAllAccounts()
    return accounts[0] ?? null
  }

  private acquireTokenInteractive = async (
    scopes: string[],
  ): Promise<AuthenticationResult> => {
    this.ensureConfigured()

    return this.clientApplication!.acquireTokenInteractive({
      scopes,
      redirectUri: REDIRECT_URI,
      openBrowser: async (url) => {
        await shell.openExternal(url)
      },
      successTemplate:
        '<h1>TeamSpy connected.</h1><p>You can close this window and return to the app.</p>',
      errorTemplate:
        '<h1>TeamSpy could not complete sign-in.</h1><p>Return to the app for the error details.</p>',
    })
  }
}
