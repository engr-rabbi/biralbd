// Admin login state. The Apps Script backend hands out a signed token after a
// successful login; it is kept in this browser's localStorage until it expires.

const KEY = 'biralbond-admin-session'
export const ADMIN_EXPIRED_EVENT = 'biralbond-admin-expired'

export type AdminSession = {
  token: string
  exp: number // ms since epoch
  user: { email: string; name: string; role: string }
}

export function getAdminSession(): AdminSession | null {
  if (typeof window === 'undefined') return null
  try {
    const raw = localStorage.getItem(KEY)
    if (!raw) return null
    const s = JSON.parse(raw) as AdminSession
    if (!s?.token || !s.exp || Date.now() > s.exp) {
      localStorage.removeItem(KEY)
      return null
    }
    return s
  } catch {
    return null
  }
}

export function saveAdminSession(s: AdminSession) {
  try { localStorage.setItem(KEY, JSON.stringify(s)) } catch { /* private mode: login lasts until reload */ }
}

export function clearAdminSession() {
  try { localStorage.removeItem(KEY) } catch { /* ignore */ }
}
