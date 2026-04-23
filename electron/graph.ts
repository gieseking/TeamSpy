import type { DirectoryPayload, DirectoryUser } from '../src/shared/types'
import { AuthManager, GRAPH_SCOPES } from './auth'

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'
const PHOTO_SIZE = 64
const USER_PAGE_SIZE = 999
const PRESENCE_BATCH_SIZE = 300
const MAILBOX_CONCURRENCY = 8

interface GraphUser {
  id: string
  displayName?: string | null
  givenName?: string | null
  surname?: string | null
  jobTitle?: string | null
  department?: string | null
  mail?: string | null
  userPrincipalName?: string | null
  officeLocation?: string | null
  companyName?: string | null
  signInActivity?: {
    lastSuccessfulSignInDateTime?: string | null
    lastSignInDateTime?: string | null
  } | null
  manager?: {
    displayName?: string | null
    mail?: string | null
    userPrincipalName?: string | null
  } | null
}

interface GraphCollection<T> {
  value: T[]
  '@odata.nextLink'?: string
}

interface Presence {
  id: string
  availability?: string | null
  activity?: string | null
  outOfOfficeSettings?: {
    isOutOfOffice?: boolean
  } | null
}

interface MailboxSettings {
  timeZone?: string | null
  workingHours?: {
    timeZone?: {
      name?: string | null
    } | null
  } | null
}

const photoCache = new Map<string, string | null>()

async function graphFetch(
  auth: AuthManager,
  endpoint: string,
  init: RequestInit = {},
  scopes?: string[],
) {
  const token = await auth.getAccessToken(scopes ?? [...GRAPH_SCOPES])
  const response = await fetch(endpoint, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init.body ? { 'Content-Type': 'application/json' } : {}),
      ...(init.headers ?? {}),
    },
  })

  if (!response.ok) {
    const body = await response.text()
    throw new Error(`${response.status} ${response.statusText}: ${body}`)
  }

  return response
}

async function graphJson<T>(
  auth: AuthManager,
  endpoint: string,
  init: RequestInit = {},
  scopes?: string[],
) {
  const response = await graphFetch(auth, endpoint, init, scopes)
  return (await response.json()) as T
}

function chunk<T>(items: T[], size: number) {
  const chunks: T[][] = []

  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size))
  }

  return chunks
}

async function mapLimit<T, R>(
  items: T[],
  concurrency: number,
  mapper: (item: T, index: number) => Promise<R>,
) {
  const results = new Array<R>(items.length)
  let nextIndex = 0

  async function worker() {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex
      nextIndex += 1
      results[currentIndex] = await mapper(items[currentIndex], currentIndex)
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(concurrency, items.length) }, () => worker()),
  )

  return results
}

function humanizePresence(presence: Presence | undefined) {
  if (!presence) {
    return {
      availability: null,
      activity: null,
      status: 'Unknown',
    }
  }

  if (presence.outOfOfficeSettings?.isOutOfOffice) {
    return {
      availability: presence.availability ?? null,
      activity: presence.activity ?? null,
      status: 'Out of office',
    }
  }

  return {
    availability: presence.availability ?? null,
    activity: presence.activity ?? null,
    status:
      presence.availability?.replace(/([a-z])([A-Z])/g, '$1 $2') ?? 'Unknown',
  }
}

function resolveLastSeen(user: GraphUser) {
  return (
    user.signInActivity?.lastSuccessfulSignInDateTime ??
    user.signInActivity?.lastSignInDateTime ??
    null
  )
}

async function listUsers(auth: AuthManager, notes: string[]) {
  const withSignInActivity =
    `${GRAPH_ROOT}/users` +
    `?$top=${USER_PAGE_SIZE}` +
    '&$select=id,displayName,givenName,surname,jobTitle,department,mail,userPrincipalName,officeLocation,companyName,signInActivity' +
    '&$expand=manager($select=displayName,mail,userPrincipalName)'

  const withoutSignInActivity =
    `${GRAPH_ROOT}/users` +
    `?$top=${USER_PAGE_SIZE}` +
    '&$select=id,displayName,givenName,surname,jobTitle,department,mail,userPrincipalName,officeLocation,companyName' +
    '&$expand=manager($select=displayName,mail,userPrincipalName)'

  const users: GraphUser[] = []
  let nextLink: string | undefined = withSignInActivity
  let includeSignInActivity = true

  while (nextLink) {
    try {
      const page: GraphCollection<GraphUser> = await graphJson(auth, nextLink)
      users.push(...page.value)
      nextLink = page['@odata.nextLink']
    } catch (error) {
      if (!includeSignInActivity) {
        throw error
      }

      includeSignInActivity = false
      nextLink = withoutSignInActivity
      users.length = 0
      notes.push(
        'Last seen is blank because sign-in activity requires extra Graph privileges in some tenants.',
      )
    }
  }

  return users
}

async function getPresenceMap(auth: AuthManager, userIds: string[]) {
  const presenceMap = new Map<string, Presence>()

  for (const ids of chunk(userIds, PRESENCE_BATCH_SIZE)) {
    const payload = await graphJson<{ value: Presence[] }>(
      auth,
      `${GRAPH_ROOT}/communications/getPresencesByUserId`,
      {
        method: 'POST',
        body: JSON.stringify({ ids }),
      },
    )

    for (const presence of payload.value) {
      presenceMap.set(presence.id, presence)
    }
  }

  return presenceMap
}

async function getMailboxMap(auth: AuthManager, users: GraphUser[], notes: string[]) {
  let warned = false

  const results = await mapLimit(users, MAILBOX_CONCURRENCY, async (user) => {
    try {
      const mailbox = await graphJson<MailboxSettings>(
        auth,
        `${GRAPH_ROOT}/users/${user.id}/mailboxSettings`,
      )

      return [user.id, mailbox] as const
    } catch {
      if (!warned) {
        warned = true
        notes.push(
          'Timezone falls back to blank when mailbox settings are unavailable for a user or tenant policy.',
        )
      }

      return [user.id, null] as const
    }
  })

  return new Map(results)
}

export async function loadDirectoryData(auth: AuthManager): Promise<DirectoryPayload> {
  const notes = [
    'Work location currently falls back to office location because Microsoft Graph does not expose the dynamic Teams work-location signal here.',
  ]

  const users = await listUsers(auth, notes)
  const presenceMap = await getPresenceMap(
    auth,
    users.map((user) => user.id),
  )
  const mailboxMap = await getMailboxMap(auth, users, notes)

  const directoryUsers: DirectoryUser[] = users.map((user) => {
    const presence = humanizePresence(presenceMap.get(user.id))
    const mailbox = mailboxMap.get(user.id)
    const email = user.mail ?? user.userPrincipalName ?? ''
    const officeLocation = user.officeLocation ?? null

    return {
      id: user.id,
      displayName: user.displayName ?? email ?? 'Unknown user',
      status: presence.status,
      availability: presence.availability,
      activity: presence.activity,
      givenName: user.givenName ?? '',
      surname: user.surname ?? '',
      jobTitle: user.jobTitle ?? '',
      department: user.department ?? '',
      email,
      userPrincipalName: user.userPrincipalName ?? email,
      reportsTo:
        user.manager?.displayName ??
        user.manager?.mail ??
        user.manager?.userPrincipalName ??
        '',
      organization: user.companyName ?? '',
      lastSeen: resolveLastSeen(user),
      timeZone:
        mailbox?.timeZone ?? mailbox?.workingHours?.timeZone?.name ?? null,
      workLocation: officeLocation,
      officeLocation,
    }
  })

  return {
    users: directoryUsers,
    loadedAt: new Date().toISOString(),
    notes,
  }
}

export async function getUserPhoto(auth: AuthManager, userId: string) {
  if (photoCache.has(userId)) {
    return photoCache.get(userId) ?? null
  }

  try {
    const response = await graphFetch(
      auth,
      `${GRAPH_ROOT}/users/${userId}/photos/${PHOTO_SIZE}x${PHOTO_SIZE}/$value`,
    )
    const contentType = response.headers.get('content-type') ?? 'image/jpeg'
    const buffer = Buffer.from(await response.arrayBuffer())
    const dataUrl = `data:${contentType};base64,${buffer.toString('base64')}`

    photoCache.set(userId, dataUrl)
    return dataUrl
  } catch {
    photoCache.set(userId, null)
    return null
  }
}
