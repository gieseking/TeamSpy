import { useEffect, useState } from 'react'
import type { DirectoryUser } from '../shared/types'

const photoCache = new Map<string, string | null>()

function getInitials(user: DirectoryUser) {
  const first = user.givenName?.[0] ?? user.displayName?.[0] ?? '?'
  const last = user.surname?.[0] ?? user.displayName?.split(' ')[1]?.[0] ?? ''
  return `${first}${last}`.trim().toUpperCase()
}

export function UserPhoto({ user }: { user: DirectoryUser }) {
  const [src, setSrc] = useState<string | null | undefined>(() =>
    photoCache.get(user.id),
  )

  useEffect(() => {
    if (photoCache.has(user.id)) {
      return
    }

    let active = true

    window.teamspy.directory
      .getPhoto(user.id)
      .then((dataUrl) => {
        photoCache.set(user.id, dataUrl)

        if (active) {
          setSrc(dataUrl)
        }
      })
      .catch(() => {
        photoCache.set(user.id, null)

        if (active) {
          setSrc(null)
        }
      })

    return () => {
      active = false
    }
  }, [user.id])

  if (src) {
    return <img className="avatar-image" src={src} alt={`${user.displayName} profile`} />
  }

  return <span className="avatar-fallback">{getInitials(user)}</span>
}
