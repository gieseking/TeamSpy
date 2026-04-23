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
import type { AuthAccount, AuthState } from '../src/shared/types'
import { AUTHORITY_TENANT, PUBLISHER_CLIENT_ID } from './publisher-config'

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

function resolveClientId() {
  return (
    process.env.TEAMSPY_CLIENT_ID?.trim() ||
    PUBLISHER_CLIENT_ID.trim() ||
    ''
  )
}

export class AuthManager {
  private readonly clientApplication: PublicClientApplication | null
  private readonly clientId: string
  private readonly authority: string

  constructor() {
    loadDesktopEnv()

    this.clientId = resolveClientId()
    this.authority = `https://login.microsoftonline.com/${AUTHORITY_TENANT}`

    if (!this.clientId) {
      this.clientApplication = null
      return
    }

    const config: Configuration = {
      auth: {
        clientId: this.clientId,
        authority: this.authority,
      },
      cache: {
        cachePlugin: new FileCachePlugin(
          path.join(app.getPath('userData'), 'msal-cache.json'),
        ),
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

  getAuthState = async (): Promise<AuthState> => {
    const account = await this.getAccount()

    return {
      configured: this.clientApplication !== null,
      signedIn: account !== null,
      account: mapAccount(account),
      missingConfiguration: this.clientId ? [] : ['clientId'],
    }
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

  private ensureConfigured() {
    if (!this.clientApplication) {
      throw new Error(
        'This TeamSpy build is missing a publisher-configured Microsoft client ID.',
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
