export type TeamsAction = 'chat' | 'call'

export interface AuthAccount {
  homeAccountId: string
  name: string
  username: string
  tenantId?: string
}

export interface AuthState {
  configured: boolean
  signedIn: boolean
  account: AuthAccount | null
  missingConfiguration: string[]
}

export interface DirectoryUser {
  id: string
  displayName: string
  status: string
  availability: string | null
  activity: string | null
  givenName: string
  surname: string
  jobTitle: string
  department: string
  email: string
  userPrincipalName: string
  reportsTo: string
  organization: string
  lastSeen: string | null
  timeZone: string | null
  workLocation: string | null
  officeLocation: string | null
}

export interface DirectoryPayload {
  users: DirectoryUser[]
  loadedAt: string
  notes: string[]
}

export interface TeamSpyDesktopApi {
  auth: {
    getState: () => Promise<AuthState>
    login: () => Promise<AuthState>
    logout: () => Promise<AuthState>
  }
  directory: {
    load: () => Promise<DirectoryPayload>
    getPhoto: (userId: string) => Promise<string | null>
  }
  teams: {
    openAction: (action: TeamsAction, email: string) => Promise<boolean>
  }
}
