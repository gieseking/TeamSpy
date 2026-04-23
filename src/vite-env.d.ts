/// <reference types="vite/client" />

import type { FriendlyFacesDesktopApi } from './shared/types'

declare global {
  interface Window {
    friendlyfaces: FriendlyFacesDesktopApi
  }
}

export {}
