// ---------------------------------------------------------------------------
// Low-level client for the Google Apps Script web app ("BiralBond API") that
// stores everything in a Google Sheet. See DEPLOY.md and
// google-apps-script/Code.gs.
//
// Every call is a POST with Content-Type text/plain. That is a "simple" CORS
// request, so the browser sends it straight away without a preflight (Apps
// Script cannot answer preflights).
// ---------------------------------------------------------------------------

// Set in the GitHub workflow from the NEXT_PUBLIC_SHEETS_API_URL secret/variable.
export const SHEETS_API_URL = (process.env.NEXT_PUBLIC_SHEETS_API_URL || '').trim()

// "" when the site lives at the domain root, "/repo-name" for GitHub project pages.
export const BASE_PATH = process.env.NEXT_PUBLIC_BASE_PATH || ''

export const isConfigured = /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(SHEETS_API_URL)

export class ApiError extends Error {
  code: string
  constructor(message: string, code = 'ERROR') {
    super(message)
    this.name = 'ApiError'
    this.code = code
  }
}

// The bridge replaces window.fetch; calls to Google must use the original one.
let realFetch: typeof fetch | null = null
export function setRealFetch(f: typeof fetch) {
  realFetch = f
}

type RpcOptions = { token?: string; timeoutMs?: number }

export async function rpc<T = any>(action: string, payload: Record<string, unknown> = {}, opts: RpcOptions = {}): Promise<T> {
  if (!isConfigured) {
    throw new ApiError(
      'NEXT_PUBLIC_SHEETS_API_URL is missing or wrong (it must look like https://script.google.com/macros/s/.../exec) - see DEPLOY.md',
      'NOT_CONFIGURED',
    )
  }
  const doFetch = realFetch ?? window.fetch.bind(window)
  const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null
  const timer = ctrl ? setTimeout(() => ctrl.abort(), opts.timeoutMs ?? 40000) : null
  let text: string
  try {
    const res = await doFetch(SHEETS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, token: opts.token, ...payload }),
      redirect: 'follow',
      cache: 'no-store',
      signal: ctrl?.signal,
    })
    text = await res.text()
  } catch (e) {
    throw new ApiError(
      e instanceof DOMException && e.name === 'AbortError'
        ? 'The Google Sheets server took too long to answer'
        : 'Cannot reach the Google Sheets server (check the Web app URL and that access is set to "Anyone")',
      'NETWORK',
    )
  } finally {
    if (timer) clearTimeout(timer)
  }
  let json: { ok?: boolean; data?: T; error?: string; code?: string }
  try {
    json = JSON.parse(text)
  } catch {
    // Google answered with an HTML page: not deployed, wrong access setting, or a script error.
    throw new ApiError('The Google Sheets server returned an unexpected page (is the Web app deployed with access "Anyone"?)', 'BAD_RESPONSE')
  }
  if (!json.ok) throw new ApiError(json.error || 'Request failed', json.code || 'ERROR')
  return json.data as T
}
