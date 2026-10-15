/* eslint-disable @typescript-eslint/no-explicit-any */
// ---------------------------------------------------------------------------
// GitHub Pages can only serve static files, so the site has no `/api/*` server.
// Instead of rewriting every component, this bridge answers the same
// `/api/...` requests (fetch calls inside the app) from the Google Sheet, via
// the Apps Script web app in google-apps-script/Code.gs, and returns exactly
// the JSON the old routes returned. The UI code is therefore unchanged.
//
// Speed: the public content is downloaded ONCE (one "bootstrap" request),
// kept in memory and in localStorage for a few minutes, and every list /
// filter / search on the site is done in the browser from that copy.
// ---------------------------------------------------------------------------
import { rpc, ApiError, BASE_PATH, isConfigured, setRealFetch } from './sheets-api'
import { MODEL_CONFIG } from './admin-config'
import { getAdminSession, saveAdminSession, clearAdminSession, ADMIN_EXPIRED_EVENT } from './admin-session'

type Row = Record<string, any>
type Ctx = { method: string; q: URLSearchParams; body: any; form: FormData | null }
type Result = { status: number; body: unknown }

const ok = (body: unknown, status = 200): Result => ({ status, body })
const fail = (error: string, status = 500): Result => ({ status, body: { error } })

// ---------- image paths ----------
const ASSET_RE = /^\/(uploads|cats)\//

// Files in /public (e.g. /cats/x.png) need the repo prefix on project pages.
function addBase(v: any): any {
  if (!BASE_PATH) return v
  if (typeof v === 'string') return ASSET_RE.test(v) ? BASE_PATH + v : v
  if (Array.isArray(v)) return v.map(addBase)
  if (v && typeof v === 'object') {
    const o: Record<string, any> = {}
    for (const k of Object.keys(v)) o[k] = addBase(v[k])
    return o
  }
  return v
}
function stripBase(v: any): any {
  if (!BASE_PATH) return v
  if (typeof v === 'string') return v.startsWith(BASE_PATH + '/') && ASSET_RE.test(v.slice(BASE_PATH.length)) ? v.slice(BASE_PATH.length) : v
  return v
}

// ---------- public data: one download, cached ----------
type PublicData = { tables: Record<string, Row[]>; counts: { members: number } }
type Stored = { at: number; data: PublicData }

const STORE_KEY = 'biralbond:public-data:v1'
const FRESH_MS = 5 * 60 * 1000 // how long a downloaded copy is trusted

let mem: Stored | null = null
let inflight: Promise<PublicData> | null = null

function readStored(): Stored | null {
  try {
    const raw = localStorage.getItem(STORE_KEY)
    if (!raw) return null
    const s = JSON.parse(raw) as Stored
    return s && s.data && s.data.tables && typeof s.at === 'number' ? s : null
  } catch {
    return null
  }
}
function writeStored(s: Stored) {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(s)) } catch { /* storage full or blocked: memory copy still works */ }
}

// Called after every change made through the site, so the next read is fresh.
export function invalidatePublic() {
  mem = null
  inflight = null
  try { localStorage.removeItem(STORE_KEY) } catch { /* ignore */ }
}

function loadPublic(): Promise<PublicData> {
  if (mem && Date.now() - mem.at < FRESH_MS) return Promise.resolve(mem.data)
  if (inflight) return inflight
  const p = (async (): Promise<PublicData> => {
    const stored = readStored()
    if (stored && Date.now() - stored.at < FRESH_MS) {
      mem = stored
      return stored.data
    }
    try {
      const data = await rpc<PublicData>('bootstrap')
      if (!data || !data.tables) throw new ApiError('Bad data from server', 'BAD_RESPONSE')
      mem = { at: Date.now(), data }
      writeStored(mem)
      return data
    } catch (e) {
      // Server unreachable: show the last copy we have instead of an empty site.
      if (stored) {
        console.error('[BiralBond] Could not refresh data, showing the last saved copy:', e)
        mem = { at: Date.now() - FRESH_MS + 30_000, data: stored.data } // retry in 30 s
        return stored.data
      }
      throw e
    }
  })()
  inflight = p
  const clear = () => { if (inflight === p) inflight = null }
  p.then(clear, clear)
  return p
}

// ---------- in-browser list / filter / search ----------
const cmp = (a: any, b: any) => {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  const sa = String(a), sb = String(b)
  return sa < sb ? -1 : sa > sb ? 1 : 0
}
function sortBy(rows: Row[], col: string, asc: boolean): Row[] {
  const withVal = rows.filter((r) => r[col] != null)
  const without = rows.filter((r) => r[col] == null)
  withVal.sort((x, y) => (asc ? 1 : -1) * cmp(x[col], y[col]))
  return [...withVal, ...without] // empty values always last
}
const has = (v: any, term: string) => String(v ?? '').toLowerCase().includes(term.toLowerCase())

type ListCfg = {
  table: string
  order?: [string, boolean] // [column, ascending]
  eqAll?: string[] // ?col=value filters ("all" = no filter)
  trueFlags?: Record<string, string> // ?param=true  ->  column = true
  search?: string[] // ?q=  matched (case-insensitive) against these columns
  extra?: (rows: Row[], q: URLSearchParams) => Row[]
}

function applyList(rows: Row[], cfg: ListCfg, q: URLSearchParams): Row[] {
  let out = rows
  for (const c of cfg.eqAll ?? []) {
    const v = q.get(c)
    if (v && v !== 'all') out = out.filter((r) => String(r[c]) === v)
  }
  for (const [param, col] of Object.entries(cfg.trueFlags ?? {})) {
    if (q.get(param) === 'true') out = out.filter((r) => r[col] === true)
  }
  const term = q.get('q')
  if (cfg.search && term) out = out.filter((r) => cfg.search!.some((c) => has(r[c], term)))
  if (cfg.extra) out = cfg.extra(out, q)
  if (cfg.order) out = sortBy(out, cfg.order[0], cfg.order[1])
  return out
}

function list(cfg: ListCfg) {
  return async ({ q }: Ctx): Promise<Result> => {
    const data = await loadPublic()
    return ok(addBase(applyList(data.tables[cfg.table] ?? [], cfg, q)))
  }
}

const group = (rows: Row[], col: string) => {
  const m = new Map<string, number>()
  for (const r of rows) {
    const k = String(r[col] ?? '').trim()
    if (k) m.set(k, (m.get(k) ?? 0) + 1)
  }
  return [...m.entries()].map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value)
}

// ---------- admin helpers ----------
async function adminCall<T = any>(action: string, payload: Record<string, unknown> = {}): Promise<T> {
  const session = getAdminSession()
  if (!session) throw new ApiError('Login required', 'AUTH')
  try {
    return await rpc<T>(action, payload, { token: session.token })
  } catch (e) {
    if (e instanceof ApiError && e.code === 'AUTH') {
      clearAdminSession()
      if (typeof window !== 'undefined') window.dispatchEvent(new Event(ADMIN_EXPIRED_EVENT))
    }
    throw e
  }
}

const sortNewest = (rows: Row[]) => sortBy(rows, 'createdAt', false)

function privateList(table: string) {
  return async (): Promise<Result> => {
    // Only admins can read these; anyone else just gets an empty list.
    if (!getAdminSession()) return ok([])
    try {
      return ok(sortNewest(await adminCall<Row[]>('adminList', { table })))
    } catch (e) {
      console.error(`[BiralBond] ${table} list error:`, e)
      return ok([])
    }
  }
}

// ---------- visitor forms ----------
function submitRoute(table: string, failMsg: string, status = 201) {
  return async ({ body }: Ctx): Promise<Result> => {
    const row = await rpc<Row>('submit', { table, data: body ?? {} }).catch((e) => {
      if (e instanceof ApiError && (e.code === 'VALIDATION' || e.code === 'RATE_LIMIT')) throw e
      console.error(`[BiralBond] ${table} submit error:`, e)
      throw new ApiError(failMsg, 'ERROR')
    })
    invalidatePublic()
    return ok(table === 'Newsletter' ? row : addBase(row), status)
  }
}

// ---------- image upload (admin) ----------
const IMAGE_TYPES = ['image/jpeg', 'image/png', 'image/webp', 'image/gif']
const MAX_SIDE = 1600

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(String(r.result).split(',')[1] || '')
    r.onerror = () => reject(new Error('read failed'))
    r.readAsDataURL(blob)
  })
}

// Shrinks big photos in the browser (keeps uploads fast and Drive small).
async function downscale(file: File): Promise<Blob | null> {
  if (typeof createImageBitmap !== 'function') return null
  const bmp = await createImageBitmap(file)
  const scale = Math.min(1, MAX_SIDE / Math.max(bmp.width, bmp.height))
  if (scale === 1 && file.size <= 800 * 1024) return null
  const canvas = document.createElement('canvas')
  canvas.width = Math.round(bmp.width * scale)
  canvas.height = Math.round(bmp.height * scale)
  canvas.getContext('2d')!.drawImage(bmp, 0, 0, canvas.width, canvas.height)
  const blob = await new Promise<Blob | null>((res) => canvas.toBlob(res, 'image/webp', 0.85))
  return blob && blob.size < file.size ? blob : null
}

async function prepareImage(file: File): Promise<{ name: string; mime: string; base64: string }> {
  if (!IMAGE_TYPES.includes(file.type)) throw new ApiError('Only image files allowed (JPEG, PNG, WebP, GIF)', 'VALIDATION')
  if (file.size > 5 * 1024 * 1024) throw new ApiError('File too large (max 5MB)', 'VALIDATION')
  let blob: Blob = file
  if (file.type !== 'image/gif') {
    try { blob = (await downscale(file)) ?? file } catch { blob = file }
  }
  return { name: file.name, mime: blob.type || file.type, base64: await blobToBase64(blob) }
}

// ---------- routes ----------
const ROUTES: Record<string, (c: Ctx) => Promise<Result>> = {
  // ----- public content (filtered in the browser) -----
  'GET /api/breeds': list({ table: 'Breed', order: ['createdAt', true] }),
  'GET /api/caretips': list({ table: 'CareTip', order: ['createdAt', true] }),
  'GET /api/feature-toggles': list({ table: 'FeatureToggle' }),
  'GET /api/testimonials': list({ table: 'Testimonial', order: ['createdAt', false] }),
  'GET /api/events': list({ table: 'Event', order: ['date', true], eqAll: ['type'] }),
  'GET /api/gallery': list({ table: 'Gallery', order: ['likes', false], eqAll: ['city'] }),
  'GET /api/articles': list({ table: 'Article', order: ['createdAt', false], eqAll: ['category'], trueFlags: { featured: 'featured' } }),
  'GET /api/lostfound': list({ table: 'LostFound', order: ['createdAt', false], eqAll: ['type'] }),
  'GET /api/community': list({ table: 'CommunityPost', order: ['createdAt', false], eqAll: ['category'] }),
  'GET /api/cat-shops': list({
    table: 'CatShop', order: ['rating', false], eqAll: ['division', 'district', 'area', 'verified'],
    search: ['shopName', 'ownerName', 'area', 'availableBreeds'],
  }),
  'GET /api/cat-food-shops': list({
    table: 'CatFoodShop', order: ['rating', false], eqAll: ['division', 'district', 'area', 'verified'],
    search: ['shopName', 'area', 'availableBrands', 'foodTypes'],
    extra: (rows, q) => {
      const brand = q.get('brand')
      return brand && brand !== 'all' ? rows.filter((r) => has(r.availableBrands, brand)) : rows
    },
  }),
  'GET /api/service-providers': list({
    table: 'ServiceProvider', order: ['rating', false], eqAll: ['type', 'division', 'district', 'area', 'verified'],
    search: ['providerName', 'businessName', 'area', 'description'],
  }),
  'GET /api/veterinary-clinics': list({
    table: 'VeterinaryClinic', order: ['rating', false], eqAll: ['division', 'district', 'area', 'verified'],
    trueFlags: { emergency: 'emergencyService' }, search: ['clinicName', 'doctorName', 'area', 'services'],
  }),

  // private lists (admins only; the server checks the token)
  'GET /api/members': privateList('Member'),
  'GET /api/contact': privateList('Contact'),
  'GET /api/newsletter': privateList('Newsletter'),

  'GET /api/comments': async ({ q }) => {
    const postId = q.get('postId')
    if (!postId) return ok([])
    return ok(await rpc<Row[]>('comments', { postId }))
  },

  'GET /api/stats': async () => {
    const { tables: t, counts } = await loadPublic()
    const n = (k: string) => (t[k] ?? []).length
    return ok({
      totals: {
        breeds: n('Breed'), catShops: n('CatShop'), catFoodShops: n('CatFoodShop'), vetClinics: n('VeterinaryClinic'),
        serviceProviders: n('ServiceProvider'), articles: n('Article'), communityPosts: n('CommunityPost'),
        events: n('Event'), gallery: n('Gallery'), members: counts?.members ?? 0, testimonials: n('Testimonial'),
        lostFound: n('LostFound'),
      },
      charts: {
        breedByCategory: group(t.Breed ?? [], 'category'),
        vetByCity: group(t.VeterinaryClinic ?? [], 'division'),
        serviceByType: group(t.ServiceProvider ?? [], 'type'),
      },
    })
  },

  'GET /api/search': async ({ q }) => {
    const term = (q.get('q') || '').trim()
    const empty = { catShops: [], catFoodShops: [], vetClinics: [], serviceProviders: [], breeds: [], articles: [], events: [], communityPosts: [], lostFound: [], total: 0 }
    if (!term) return ok(empty)
    const { tables: t } = await loadPublic()
    const find = (table: string, cols: string[]) =>
      (t[table] ?? []).filter((r) => cols.some((c) => has(r[c], term))).slice(0, 5)
    const res = {
      catShops: find('CatShop', ['shopName', 'area', 'availableBreeds', 'division']),
      catFoodShops: find('CatFoodShop', ['shopName', 'area', 'availableBrands', 'division']),
      vetClinics: find('VeterinaryClinic', ['clinicName', 'doctorName', 'area', 'division']),
      serviceProviders: find('ServiceProvider', ['providerName', 'area', 'type', 'division']),
      breeds: find('Breed', ['name', 'bnName', 'temperament']),
      articles: find('Article', ['title', 'excerpt', 'category']),
      events: find('Event', ['title', 'city', 'type']),
      communityPosts: find('CommunityPost', ['title', 'body', 'author']),
      lostFound: find('LostFound', ['catName', 'breed', 'city']),
    }
    const total = Object.values(res).reduce((s, a) => s + a.length, 0)
    return ok(addBase({ ...res, total }))
  },

  // ----- visitor forms -----
  'POST /api/newsletter': submitRoute('Newsletter', 'Failed to subscribe'),
  'POST /api/members': async (c) => submitRoute('Member', 'Failed to register member')({
    ...c, body: { ...c.body, city: c.body?.city || 'Dhaka', plan: c.body?.plan || 'Free', cats: Number(c.body?.cats) || 0 },
  }),
  'POST /api/contact': async (c) => submitRoute('Contact', 'Failed to send message')({
    ...c, body: { ...c.body, subject: c.body?.subject || 'General enquiry' },
  }),
  'POST /api/community': submitRoute('CommunityPost', 'Failed to create post'),
  'POST /api/lostfound': submitRoute('LostFound', 'Failed to create listing'),
  'POST /api/comments': async (c) => {
    if (!c.body?.postId || !c.body?.author || !c.body?.body) return fail('Missing fields', 400)
    return submitRoute('Comment', 'Failed to add comment')(c)
  },

  // ----- AI (Google Gemini, called by the Apps Script) -----
  'POST /api/cat-assistant': async ({ body }) => {
    const messages = body?.messages
    if (!Array.isArray(messages) || messages.length === 0) return fail('Messages required', 400)
    try {
      const data = await rpc<{ reply?: string }>('chat', { messages: messages.slice(-12) }, { timeoutMs: 60000 })
      if (!data?.reply) throw new Error('No reply')
      return ok({ reply: data.reply })
    } catch (e) {
      // Open the browser console (F12) to see why: e.g. GEMINI_API_KEY missing, quota, not redeployed.
      console.error('[Miau] AI request failed:', e)
      return fail('Miau is taking a quick catnap. Please try again in a moment.')
    }
  },

  'POST /api/cat-names': async ({ body }) => {
    const fallback = ['Misti', 'Bagha', 'Tuntuni', 'Miau', 'Simba', 'Baghini', 'Kalo', 'Bilai']
    try {
      const data = await rpc<{ names?: string[] }>('names', { vibe: body?.vibe, gender: body?.gender, breed: body?.breed }, { timeoutMs: 60000 })
      if (!Array.isArray(data?.names) || data.names.length === 0) throw new Error('No names')
      return ok({ names: data.names.slice(0, 8) })
    } catch (e) {
      console.error('[Miau] name generator AI request failed:', e)
      return ok({ names: fallback })
    }
  },

  // ----- admin: login -----
  // The admin e-mail and password are Script Properties in your Apps Script
  // project (DEPLOY.md, step 4). The server hands out a signed token.
  'GET /api/admin/login': async () => ok({ hasAdmin: true, count: 1 }),

  'POST /api/admin/login': async ({ body }) => {
    const { email, password } = body || {}
    if (!email || !password) return fail('ইমেইল ও পাসওয়ার্ড প্রয়োজন', 400)
    try {
      const d = await rpc<{ token: string; exp: number; user: any }>('login', { email: String(email).trim(), password: String(password) })
      saveAdminSession({ token: d.token, exp: d.exp, user: d.user })
      return ok({ ok: true, user: d.user })
    } catch (e) {
      if (e instanceof ApiError) {
        if (e.code === 'BAD_LOGIN') return fail(e.message, 401)
        if (e.code === 'RATE_LIMIT') return fail(e.message, 429)
        if (e.code === 'NOT_CONFIGURED' && e.message.includes('ADMIN_')) return fail(e.message, 500)
      }
      console.error('[BiralBond] admin login error:', e)
      return fail('লগইন ব্যর্থ। ওয়েব অ্যাপ URL ও ডিপ্লয় ঠিক আছে কিনা দেখুন (DEPLOY.md)।')
    }
  },

  'POST /api/admin/register': async () =>
    fail('অ্যাডমিন ইমেইল-পাসওয়ার্ড Apps Script-এর Script properties থেকে সেট করতে হয়। DEPLOY.md দেখুন।', 403),

  'POST /api/admin/reset': async () => fail('এই সুবিধা নিরাপত্তার কারণে বন্ধ করা হয়েছে', 403),

  // ----- home page content (public read, admin write) -----
  'GET /api/admin/site-content': async ({ q }) => {
    const rows = (await loadPublic()).tables.SiteContent ?? []
    const section = q.get('section')
    if (section) return ok(addBase(rows.find((r) => r.section === section) || {}))
    return ok(addBase(sortBy(rows, 'section', true)))
  },

  'PATCH /api/admin/site-content': async ({ body }) => {
    const { section, ...rest } = body || {}
    if (!section) return fail('Section required', 400)
    const allowed = ['eyebrow', 'title', 'titleAccent', 'titleEnd', 'description', 'buttonText', 'buttonLink', 'buttonText2', 'buttonLink2', 'image']
    const data: Record<string, unknown> = {}
    for (const k of allowed) {
      if (rest[k] === undefined) continue
      data[k] = k === 'image' ? stripBase(rest[k]) : rest[k]
    }
    const row = await adminCall<Row>('siteContentSave', { section, data })
    invalidatePublic()
    return ok(addBase(row))
  },

  // ----- admin: dashboard, CRUD, upload (need an admin login) -----
  'GET /api/admin/stats': async () => {
    const d = await adminCall<any>('adminStats')
    return ok({ ...d, recent: { posts: addBase(d.recent?.posts ?? []), contacts: d.recent?.contacts ?? [] } })
  },

  'GET /api/admin/crud': async ({ q }) => {
    const model = q.get('model')
    if (!model || !MODEL_CONFIG[model]) return fail('Invalid model', 400)
    return ok({ fields: MODEL_CONFIG[model].fields })
  },

  'POST /api/admin/crud': async ({ q, body }) => {
    const model = q.get('model')
    if (!model || !MODEL_CONFIG[model]) return fail('Invalid model', 400)
    const cfg = MODEL_CONFIG[model]
    const data: Record<string, unknown> = {}
    for (const f of cfg.fields) {
      const v = body[f.name]
      if (v !== undefined && v !== '') {
        if (f.type === 'number') data[f.name] = Number(v)
        else if (f.type === 'boolean') data[f.name] = v === true || v === 'true'
        else data[f.name] = stripBase(v)
      } else if (f.type === 'boolean') data[f.name] = false
    }
    const row = await adminCall<Row>('adminCreate', { table: cfg.table, data })
    invalidatePublic()
    return ok(addBase(row), 201)
  },

  'PATCH /api/admin/crud': async ({ q, body }) => {
    const model = q.get('model')
    const id = q.get('id')
    if (!model || !id || !MODEL_CONFIG[model]) return fail('Invalid model or id', 400)
    const cfg = MODEL_CONFIG[model]
    const data: Record<string, unknown> = {}
    for (const f of cfg.fields) {
      const v = body[f.name]
      if (v === undefined) continue
      if (f.type === 'number') data[f.name] = Number(v)
      else if (f.type === 'boolean') data[f.name] = v === true || v === 'true'
      else data[f.name] = stripBase(v) || null
    }
    const row = await adminCall<Row>('adminUpdate', { table: cfg.table, id, data })
    invalidatePublic()
    return ok(addBase(row))
  },

  'DELETE /api/admin/crud': async ({ q }) => {
    const model = q.get('model')
    const id = q.get('id')
    if (!model || !id || !MODEL_CONFIG[model]) return fail('Invalid model or id', 400)
    await adminCall('adminDelete', { table: MODEL_CONFIG[model].table, id })
    invalidatePublic()
    return ok({ ok: true })
  },

  'POST /api/admin/upload': async ({ form }) => {
    const file = form?.get('file') as File | null
    if (!file) return fail('No file provided', 400)
    const img = await prepareImage(file)
    const d = await adminCall<{ url: string; filename: string }>('upload', img)
    return ok({ url: d.url, filename: d.filename })
  },
}

// Public GET endpoints whose success body is always an array. If the server
// cannot be reached we answer with an empty list, so the page still opens
// (components call .map()/.filter() on the result and would otherwise crash).
const SAFE_EMPTY_LIST = new Set([
  '/api/breeds', '/api/caretips', '/api/feature-toggles', '/api/testimonials',
  '/api/events', '/api/gallery', '/api/articles', '/api/lostfound',
  '/api/community', '/api/cat-shops', '/api/cat-food-shops',
  '/api/service-providers', '/api/veterinary-clinics', '/api/comments',
])

const EMPTY_SEARCH = { catShops: [], catFoodShops: [], vetClinics: [], serviceProviders: [], breeds: [], articles: [], events: [], communityPosts: [], lostFound: [], total: 0 }

function errorResult(e: unknown): Result {
  if (e instanceof ApiError) {
    const status = e.code === 'AUTH' ? 401 : e.code === 'VALIDATION' ? 400 : e.code === 'NOT_FOUND' ? 404 : e.code === 'RATE_LIMIT' ? 429 : 500
    return fail(e.message, status)
  }
  return fail('Failed')
}

// ---------- fetch interception ----------
function parseTarget(input: RequestInfo | URL): { path: string; search: string } | null {
  const raw = typeof input === 'string' ? input : input instanceof URL ? input.href : (input as Request).url
  let u: URL
  try { u = new URL(raw, window.location.href) } catch { return null }
  if (u.origin !== window.location.origin) return null
  let p = u.pathname
  if (BASE_PATH && p.startsWith(BASE_PATH + '/')) p = p.slice(BASE_PATH.length)
  if (!p.startsWith('/api/')) return null
  return { path: p.replace(/\/+$/, ''), search: u.search }
}

async function readBody(input: RequestInfo | URL, init?: RequestInit): Promise<{ body: any; form: FormData | null }> {
  let raw: any = init?.body
  if (raw === undefined && typeof input === 'object' && !(input instanceof URL)) {
    try { raw = await (input as Request).clone().text() } catch { raw = undefined }
  }
  if (typeof FormData !== 'undefined' && raw instanceof FormData) return { body: {}, form: raw }
  if (typeof raw === 'string' && raw) {
    try { return { body: JSON.parse(raw), form: null } } catch { /* not JSON */ }
  }
  return { body: {}, form: null }
}

export function installApiBridge() {
  if (typeof window === 'undefined') return
  const w = window as any
  if (w.__biralbondApiBridge) return
  w.__biralbondApiBridge = true

  if (!isConfigured) {
    console.error('[BiralBond] NEXT_PUBLIC_SHEETS_API_URL is not set (or is not a ".../exec" web app URL) - see DEPLOY.md')
  }

  const original = window.fetch.bind(window)
  setRealFetch(original)

  // Start downloading the public data right away, before any section asks for it.
  if (isConfigured) loadPublic().catch((e) => console.error('[BiralBond] could not load data:', e))

  window.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
    const target = parseTarget(input)
    if (!target) return original(input, init)

    const method = (init?.method || (typeof input === 'object' && !(input instanceof URL) ? (input as Request).method : 'GET') || 'GET').toUpperCase()
    const handler = ROUTES[`${method} ${target.path}`]
    let result: Result
    if (!handler) {
      result = fail('Not found', 404)
    } else {
      try {
        const { body, form } = await readBody(input, init)
        result = await handler({ method, q: new URLSearchParams(target.search), body, form })
      } catch (e) {
        console.error(`[BiralBond] ${method} ${target.path} failed:`, e)
        result = errorResult(e)
      }
    }

    if (method === 'GET' && result.status >= 400) {
      if (SAFE_EMPTY_LIST.has(target.path)) {
        return new Response('[]', { status: 200, headers: { 'Content-Type': 'application/json', 'x-biralbond-degraded': '1' } })
      }
      if (target.path === '/api/search') {
        return new Response(JSON.stringify(EMPTY_SEARCH), { status: 200, headers: { 'Content-Type': 'application/json', 'x-biralbond-degraded': '1' } })
      }
    }
    return new Response(JSON.stringify(result.body), {
      status: result.status,
      headers: { 'Content-Type': 'application/json' },
    })
  }
}
