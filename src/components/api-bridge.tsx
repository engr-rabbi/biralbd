'use client'

import { installApiBridge } from '@/lib/api-bridge'

// Installed at module load (before any component's effects run) so every
// fetch('/api/...') in the app is answered from the Google Sheet.
installApiBridge()

export function ApiBridge() {
  return null
}
