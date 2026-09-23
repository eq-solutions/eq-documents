/**
 * brand.normalise — the ONE mapping from a tenant's raw brand record to a
 * TenantBrandKit. The canonical RPC `eq_get_tenant_brand_kit` applies the
 * same rules in SQL; this is the client/server JS copy so a consumer that
 * already holds the raw `organisations.branding` jsonb (eq-field's
 * `__TENANT_BRANDING_CANON__`, eq-service's JWT claims) never needs a second
 * network call, and so the two can be tested against each other.
 *
 * Rules:
 *  - Palette roles are validated individually; a bad or missing role falls
 *    back to NEUTRAL_BRAND_KIT's grey for that role and marks the kit
 *    incomplete. Never to another tenant's value.
 *  - Structured `branding.logos.*` (with measured dimensions) wins. Legacy
 *    keys (`hubLogo` URL, `gateLogo` / `sidebarLogoHtml` HTML strings,
 *    `gateLogoDark` URL) are recognised and their URL extracted, but a URL
 *    without dimensions cannot become a LogoAsset — it is returned as
 *    `legacyLogos` for the caller to measure (see `withMeasuredLogos`).
 *  - Fonts and legal come from `branding.fonts` / `branding.legal` when
 *    present; otherwise neutral (system fonts, empty footer).
 *  - Auth-gate flags (`rememberMeDays`, `coreOnly`, …) are ignored: not brand.
 */
import {
  BRAND_KIT_VERSION,
  DOC_BODY_SAFE_FONTS,
  NEUTRAL_BRAND_KIT,
  toHex6,
  type BrandFonts,
  type BrandLegal,
  type BrandPalette,
  type LogoAsset,
  type TenantBrandKit,
} from '@eq-solutions/contracts'

export interface TenantIdentity {
  id: string
  slug: string
  /** Display name. Falls back to slug. */
  name?: string | null
  /** Legal entity name. Falls back to `name`. */
  legalName?: string | null
}

/** Raw `public.organisations.branding` jsonb, as stored. Loosely typed on purpose. */
export type RawBranding = Record<string, unknown> | null | undefined

export interface LegacyLogos {
  light?: string
  dark?: string
}

export interface NormaliseResult {
  kit: TenantBrandKit
  /** Logo URLs found only in legacy keys, without measured dimensions. Measure and merge with withMeasuredLogos(). */
  legacyLogos: LegacyLogos
  /** Which palette roles were filled from the neutral kit. Empty when the palette was complete. */
  neutralised: Array<keyof BrandPalette>
  /** True when heading/body/docBody all came from the tenant record. */
  fontsFromTenant: boolean
}

const LOGO_MIMES: LogoAsset['mime'][] = ['image/png', 'image/svg+xml', 'image/webp', 'image/jpeg']

function asRecord(v: unknown): Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v) ? (v as Record<string, unknown>) : {}
}

function asString(v: unknown): string | undefined {
  return typeof v === 'string' && v.trim() !== '' ? v.trim() : undefined
}

/** Pull the first `src="…"` out of a stored HTML snippet, or return the input if it is already a bare URL. */
export function extractImageUrl(v: unknown): string | undefined {
  const s = asString(v)
  if (!s) return undefined
  if (!s.includes('<')) return s
  const m = s.match(/<img[^>]*\ssrc\s*=\s*["']([^"']+)["']/i)
  return m?.[1]
}

/** Infer a LogoAsset mime from a URL's extension. Undefined when unknown. */
export function mimeFromUrl(url: string): LogoAsset['mime'] | undefined {
  const clean = url.split('?')[0].split('#')[0].toLowerCase()
  if (clean.endsWith('.png')) return 'image/png'
  if (clean.endsWith('.svg')) return 'image/svg+xml'
  if (clean.endsWith('.webp')) return 'image/webp'
  if (clean.endsWith('.jpg') || clean.endsWith('.jpeg')) return 'image/jpeg'
  if (clean.startsWith('data:image/png')) return 'image/png'
  if (clean.startsWith('data:image/svg+xml')) return 'image/svg+xml'
  return undefined
}

function toLogoAsset(v: unknown): LogoAsset | undefined {
  const o = asRecord(v)
  const url = asString(o.url)
  const widthPx = typeof o.widthPx === 'number' ? o.widthPx : Number(o.widthPx)
  const heightPx = typeof o.heightPx === 'number' ? o.heightPx : Number(o.heightPx)
  const mime = (asString(o.mime) as LogoAsset['mime'] | undefined) ?? (url ? mimeFromUrl(url) : undefined)
  if (!url || !(widthPx > 0) || !(heightPx > 0) || !mime || !LOGO_MIMES.includes(mime)) return undefined
  return { url, widthPx, heightPx, mime }
}

function normalisePalette(raw: unknown): { palette: BrandPalette; neutralised: Array<keyof BrandPalette> } {
  const o = asRecord(raw)
  const neutralised: Array<keyof BrandPalette> = []
  const pick = (k: 'primary' | 'deep' | 'ice' | 'ink'): string => {
    const hex = toHex6(o[k])
    if (hex) return hex
    neutralised.push(k)
    return NEUTRAL_BRAND_KIT.palette[k]
  }
  const palette: BrandPalette = { primary: pick('primary'), deep: pick('deep'), ice: pick('ice'), ink: pick('ink') }
  const accent = toHex6(o.accent)
  if (accent) palette.accent = accent
  return { palette, neutralised }
}

function normaliseFonts(raw: unknown): { fonts: BrandFonts; fromTenant: boolean } {
  const o = asRecord(raw)
  const heading = asString(o.heading)
  const body = asString(o.body)
  const docBodyRaw = asString(o.docBody)
  const docBody = docBodyRaw && (DOC_BODY_SAFE_FONTS as readonly string[]).includes(docBodyRaw) ? docBodyRaw : undefined
  const fromTenant = Boolean(heading && body && docBody)
  return {
    fonts: {
      heading: heading ?? NEUTRAL_BRAND_KIT.fonts.heading,
      body: body ?? heading ?? NEUTRAL_BRAND_KIT.fonts.body,
      docBody: docBody ?? NEUTRAL_BRAND_KIT.fonts.docBody,
    },
    fromTenant,
  }
}

function normaliseLegal(raw: unknown): BrandLegal {
  const o = asRecord(raw)
  const legal: BrandLegal = {}
  for (const k of ['abn', 'address', 'phone', 'email', 'web'] as const) {
    const v = asString(o[k])
    if (v) legal[k] = v
  }
  return legal
}

/**
 * Normalise a tenant's raw branding record into a kit. Pure and synchronous.
 * `complete` is true only when all four palette roles, a light logo with
 * dimensions, and tenant-set fonts were present.
 */
export function normalise(tenant: TenantIdentity, branding: RawBranding): NormaliseResult {
  const b = asRecord(branding)
  const { palette, neutralised } = normalisePalette(b.palette)
  const { fonts, fromTenant: fontsFromTenant } = normaliseFonts(b.fonts)
  const legal = normaliseLegal(b.legal)

  const structured = asRecord(b.logos)
  const logos: TenantBrandKit['logos'] = {}
  const light = toLogoAsset(structured.light)
  const dark = toLogoAsset(structured.dark)
  const mark = toLogoAsset(structured.mark)
  if (light) logos.light = light
  if (dark) logos.dark = dark
  if (mark) logos.mark = mark

  const legacyLogos: LegacyLogos = {}
  if (!light) {
    // Document/gate logo first (it is the print-quality one), then hub/nav logo.
    const url = extractImageUrl(b.gateLogo) ?? extractImageUrl(b.sidebarLogoHtml) ?? extractImageUrl(b.hubLogo)
    if (url) legacyLogos.light = url
  }
  if (!dark) {
    const url = extractImageUrl(b.gateLogoDark)
    if (url) legacyLogos.dark = url
  }

  // Mirrors eq_brand_kit_from_branding() on jvkn exactly: display name is the
  // organisation's name (branding.orgName is an auth-gate label, e.g.
  // "EQ Solves — Field", not a document masthead); legal name comes from
  // branding.legal.legalName, else the caller's legalName, else the name.
  const displayName = asString(tenant.name) ?? tenant.slug
  const legalName = asString(asRecord(b.legal).legalName) ?? asString(tenant.legalName) ?? displayName

  const kit: TenantBrandKit = {
    kitVersion: BRAND_KIT_VERSION,
    tenant: { id: tenant.id, slug: tenant.slug, legalName, displayName },
    palette,
    logos,
    fonts,
    legal,
    policy: { flat: true },
    complete: neutralised.length === 0 && Boolean(light) && fontsFromTenant,
  }
  return { kit, legacyLogos, neutralised, fontsFromTenant }
}

export type MeasureFn = (url: string) => Promise<{ widthPx: number; heightPx: number } | null>

/**
 * Fill `logos.light` / `logos.dark` from legacy URLs using a caller-supplied
 * measurer (eq-field: an <img> + canvas; eq-shell functions: IHDR read; the
 * upload function: sharp/IHDR at upload time). Returns a new kit; the input
 * is not mutated. `complete` is recomputed.
 */
export async function withMeasuredLogos(result: NormaliseResult, measure: MeasureFn): Promise<TenantBrandKit> {
  const logos = { ...result.kit.logos }
  for (const role of ['light', 'dark'] as const) {
    const url = result.legacyLogos[role]
    if (!url || logos[role]) continue
    const mime = mimeFromUrl(url)
    if (!mime) continue
    const dims = await measure(url).catch(() => null)
    if (dims && dims.widthPx > 0 && dims.heightPx > 0) logos[role] = { url, widthPx: dims.widthPx, heightPx: dims.heightPx, mime }
  }
  const complete = result.neutralised.length === 0 && Boolean(logos.light) && result.fontsFromTenant
  return { ...result.kit, logos, complete }
}

/**
 * Adapter for the Shell→Service handoff JWT claims (brand_color + optional
 * brand_deep/ice/ink/doc_logo_url/doc_logo_dark_url). Produces the same
 * NormaliseResult shape so eq-service can keep the JWT as its transport.
 */
export function fromHandoffClaims(
  tenant: TenantIdentity,
  claims: Record<string, unknown>,
): NormaliseResult {
  const branding: Record<string, unknown> = {
    palette: { primary: claims.brand_color, deep: claims.brand_deep, ice: claims.brand_ice, ink: claims.brand_ink },
    gateLogo: claims.brand_doc_logo_url ?? claims.brand_logo_url,
    gateLogoDark: claims.brand_doc_logo_dark_url,
    fonts: claims.brand_fonts,
    legal: claims.brand_legal,
  }
  return normalise(tenant, branding)
}
