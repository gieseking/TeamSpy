/// <reference types="vite/client" />

import type { TeamSpyDesktopApi } from './shared/types'

declare global {
  interface Window {
    teamspy: TeamSpyDesktopApi
  }
}

export {}
