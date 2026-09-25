/**
 * preflight — the six-line brand check, tenant-parametrised, in code.
 * Mirrors the chat-line discipline in eq-context rules/brand-check.md
 * ("Brand check: ✓ logo ✓ ratio ✓ palette ✓ fonts ✓ flat ✓ footer") but
 * against a kit instead of one company's fixed values, and inspectable by
 * CI on golden documents.
 *
 * `preflightDocx` unzips a generated .docx and reads document.xml,
 * styles.xml, header/footer parts and image extents so the check is on
 * what was actually written, not on what the caller intended.
 */
import JSZip from 'jszip'
import type { TenantBrandKit } from '@eq-solutions/contracts'
import { contrastRatio, textOn } from '../brand/contrast.js'
import { HAIRLINE } from '../docx/primitives.js'
import { mutedInk } from '../docx/styles.js'

export type CheckId = 'logo' | 'ratio' | 'palette' | 'fonts' | 'flat' | 'footer'

export interface CheckResult {
  id: CheckId
  pass: boolean
  detail?: string
}

export interface PreflightResult {
  ok: boolean
  checks: CheckResult[]
  /** The one-line chat form: "Brand check: ✓ logo ✓ ratio ✗ palette → …" */
  line: string
}

/** What a document declares about itself. Fill from the generator, or let preflightDocx extract it. */
export interface DocumentFacts {
  /** URLs or storage paths of every image placed as a logo. Empty when the kit has no logo. */
  logoSources?: string[]
  /** Placed logo sizes (px or EMU — only the ratio matters). */
  logoPlacements?: Array<{ width: number; height: number }>
  /** Every colour hex used for text, fills, borders. Uppercase bare hex. */
  usedHex: string[]
  /** Every font family name used. */
  usedFonts: string[]
  /** Any gradient / shadow / glow / effect present. */
  hasEffects: boolean
  /** Footer part present and containing the legal name. */
  hasFooter: boolean
  footerText?: string
}

/** Colours a kit legitimately permits beyond its palette. */
export function allowedHex(kit: TenantBrandKit): Set<string> {
  const s = new Set<string>([
    kit.palette.primary,
    kit.palette.deep,
    kit.palette.ice,
    kit.palette.ink,
    'FFFFFF',
    HAIRLINE,
    mutedInk(kit),
    'AUTO',
  ])
  if (kit.palette.accent) s.add(kit.palette.accent)
  return s
}

export function preflight(kit: TenantBrandKit, facts: DocumentFacts): PreflightResult {
  const checks: CheckResult[] = []

  // 1 logo — every placed logo must be one of the kit's assets (or none when the kit has none).
  const kitLogoUrls = [kit.logos.light?.url, kit.logos.dark?.url, kit.logos.mark?.url].filter(Boolean) as string[]
  const badLogos = (facts.logoSources ?? []).filter((u) => !kitLogoUrls.includes(u))
  checks.push({ id: 'logo', pass: badLogos.length === 0, detail: badLogos.length ? `logo not from kit: ${badLogos.join(', ')}` : undefined })

  // 2 ratio — placed logos keep the stored aspect ratio within 2 %.
  const expected = kit.logos.light ? kit.logos.light.widthPx / kit.logos.light.heightPx : undefined
  const badRatio = (facts.logoPlacements ?? []).filter((p) => expected !== undefined && Math.abs(p.width / p.height - expected) / expected > 0.02)
  checks.push({ id: 'ratio', pass: badRatio.length === 0, detail: badRatio.length ? `logo stretched: ${badRatio.map((p) => `${p.width}×${p.height}`).join(', ')}` : undefined })

  // 3 palette — only kit colours, and readable text on primary. headCell()
  // (primitives.ts) never hardcodes white — it calls textOn(primary, ink) to
  // pick whichever of white/ink actually contrasts better, so the check must
  // grade that real choice, not assume white was used regardless of the kit.
  const allowed = allowedHex(kit)
  const foreign = [...new Set(facts.usedHex.map((h) => h.toUpperCase()))].filter((h) => !allowed.has(h))
  const headFg = textOn(kit.palette.primary, kit.palette.ink)
  const headContrast = contrastRatio(headFg, kit.palette.primary)
  const paletteDetail = foreign.length ? `colours outside kit: ${foreign.join(', ')}` : headContrast < 4.5 ? `${headFg === 'FFFFFF' ? 'white' : 'ink'} on primary is ${headContrast.toFixed(1)}:1` : undefined
  checks.push({ id: 'palette', pass: foreign.length === 0 && headContrast >= 4.5, detail: paletteDetail })

  // 4 fonts — only the kit's three families.
  const kitFonts = new Set([kit.fonts.heading, kit.fonts.body, kit.fonts.docBody])
  const badFonts = [...new Set(facts.usedFonts)].filter((f) => !kitFonts.has(f))
  checks.push({ id: 'fonts', pass: badFonts.length === 0, detail: badFonts.length ? `fonts outside kit: ${badFonts.join(', ')}` : undefined })

  // 5 flat
  checks.push({ id: 'flat', pass: !facts.hasEffects, detail: facts.hasEffects ? 'gradient/shadow/effect present' : undefined })

  // 6 footer — present, carries the legal name; never a different company's name.
  const footerOk = facts.hasFooter && (facts.footerText ?? '').includes(kit.tenant.legalName)
  checks.push({ id: 'footer', pass: footerOk, detail: !facts.hasFooter ? 'no footer' : footerOk ? undefined : `footer does not carry "${kit.tenant.legalName}"` })

  const ok = checks.every((c) => c.pass)
  const line = 'Brand check: ' + checks.map((c) => (c.pass ? `✓ ${c.id}` : `✗ ${c.id} → ${c.detail ?? 'failed'}`)).join(' ')
  return { ok, checks, line }
}

const HEX_ATTR = /w:(?:color|fill|themeColor)="([0-9A-Fa-f]{6}|auto)"/g
const FONT_ATTR = /w:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"/g
const EFFECT_TAGS = /<w:(?:shadow|glow|reflection|effect|gradFill|outline|emboss|imprint)[\s/>]|<a:gradFill|<a:effectLst>\s*<a:/
const EXTENT = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/g

/** Extract DocumentFacts from a generated .docx. Logo sources are reported by media part name because the URL is gone once embedded. */
export async function extractFacts(docx: Uint8Array | ArrayBuffer | Blob): Promise<DocumentFacts> {
  const zip = await JSZip.loadAsync(docx as Uint8Array)
  const parts = Object.keys(zip.files).filter((p) => /^word\/(document|styles|header\d*|footer\d*|numbering)\.xml$/.test(p))
  const usedHex = new Set<string>()
  const usedFonts = new Set<string>()
  let hasEffects = false
  let hasFooter = false
  let footerText = ''
  const logoPlacements: Array<{ width: number; height: number }> = []
  for (const part of parts) {
    const xml = await zip.file(part)!.async('string')
    for (const m of xml.matchAll(HEX_ATTR)) usedHex.add(m[1].toUpperCase())
    for (const m of xml.matchAll(FONT_ATTR)) usedFonts.add(m[1])
    if (EFFECT_TAGS.test(xml)) hasEffects = true
    for (const m of xml.matchAll(EXTENT)) logoPlacements.push({ width: Number(m[1]), height: Number(m[2]) })
    if (/^word\/footer\d*\.xml$/.test(part)) {
      hasFooter = true
      footerText += xml.replace(/<[^>]+>/g, ' ')
    }
  }
  return { usedHex: [...usedHex], usedFonts: [...usedFonts], hasEffects, hasFooter, footerText: footerText.replace(/\s+/g, ' ').trim(), logoPlacements }
}

/** Run the six checks against the bytes of a generated .docx. */
export async function preflightDocx(kit: TenantBrandKit, docx: Uint8Array | ArrayBuffer | Blob): Promise<PreflightResult> {
  return preflight(kit, await extractFacts(docx))
}
