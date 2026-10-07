/**
 * docx.styles — named Word styles generated from a TenantBrandKit.
 *
 * This is the single biggest visual-quality lever the suite lacked: every
 * current generator writes inline run properties, so a tenant opening the
 * file sees "Normal" everywhere and cannot restyle. With these, the styles
 * gallery shows DocTitle / DocH1 / DocBody… and one edit restyles the
 * document. Tenant-neutral ids on purpose — no "SKS…" / "EQ…" prefixes.
 *
 * Sizes are half-points (docx convention). Colours are bare hex from the kit.
 */
import type { IStylesOptions } from 'docx'
import { AlignmentType, BorderStyle } from 'docx'
import { DOC_BODY_SAFE_FONTS } from '@eq-solutions/contracts'
import { luminance } from '../brand/contrast.js'
import type { TenantBrandKit } from '@eq-solutions/contracts'

/**
 * Fonts every Windows/Office install ships with, plus DOC_BODY_SAFE_FONTS
 * (the same allow-list @eq-solutions/contracts already validates
 * `fonts.docBody` against) and the fonts EQ/SKS tenants already use today.
 * `fonts.heading` carries no such validation upstream — this is the actual
 * gap the guard below closes.
 */
const SAFE_FONTS = new Set<string>([...DOC_BODY_SAFE_FONTS, 'Segoe UI', 'Georgia', 'Times New Roman', 'Roboto'])

/** Word has no font-family fallback list — `w:rFonts` names exactly one face per script range (ascii/hAnsi/eastAsia/cs), so an unrecognised font can't degrade to a chain the way a CSS font-family stack would. */
export const FALLBACK_FONT = 'Arial'

/**
 * Font fallback safety net: a font on SAFE_FONTS renders exactly as
 * requested (unchanged — no behaviour change for a tenant already using a
 * known-safe font). Anything else still renders as the tenant asked for
 * ascii/hAnsi text — never silently replaced — but pins the eastAsia/cs
 * faces to FALLBACK_FONT, the one part of rFonts that can carry a different,
 * guaranteed-present name, instead of leaving those ranges to whatever
 * Word's own per-install substitution table decides.
 */
function docFont(font: string): string | { ascii: string; hAnsi: string; cs: string; eastAsia: string } {
  return SAFE_FONTS.has(font) ? font : { ascii: font, hAnsi: font, cs: FALLBACK_FONT, eastAsia: FALLBACK_FONT }
}

export const DOC_STYLE_IDS = [
  'DocTitle',
  'DocSubtitle',
  'DocH1',
  'DocH2',
  'DocH3',
  'DocBody',
  'DocSmall',
  'DocTableHead',
  'DocTableCell',
  'DocFooter',
] as const
export type DocStyleId = (typeof DOC_STYLE_IDS)[number]

/** Accent colour: explicit accent role, else deep. */
export function accentOf(kit: TenantBrandKit): string {
  return kit.palette.accent ?? kit.palette.deep
}

/** 60 % ink blended towards white — for captions/metadata without inventing a new brand colour. */
export function mutedInk(kit: TenantBrandKit): string {
  const ink = kit.palette.ink
  const mix = (i: number) => Math.round(parseInt(ink.slice(i, i + 2), 16) * 0.6 + 255 * 0.4)
  return [0, 2, 4].map((i) => mix(i).toString(16).padStart(2, '0')).join('').toUpperCase()
}

/**
 * Light surface fill for zebra rows, label columns and stat tiles. Normally the
 * kit's `ice`; but a tenant's `ice` can be a mid-tone (SKS's renders as a heavy
 * grey-lavender slab), so when it is too dark to read as a soft surface it is
 * replaced by the primary blended 90 % towards white — the same idea as
 * `tableHeadFill`'s fallback, and still a pure function of the kit.
 */
export function softFill(kit: TenantBrandKit): string {
  if (luminance(kit.palette.ice) >= SOFT_FILL_LUMINANCE_MIN) return kit.palette.ice
  const p = kit.palette.primary
  const mix = (i: number) => Math.round(parseInt(p.slice(i, i + 2), 16) * 0.1 + 255 * 0.9)
  return [0, 2, 4].map((i) => mix(i).toString(16).padStart(2, '0')).join('').toUpperCase()
}

/** Below this relative luminance a palette `ice` reads as a grey slab, not a soft surface. */
const SOFT_FILL_LUMINANCE_MIN = 0.7

/** Build the `styles` option for `new Document({ styles })`. */
export function docxStyles(kit: TenantBrandKit): IStylesOptions {
  const { primary, deep, ink } = kit.palette
  const accent = accentOf(kit)
  const heading = docFont(kit.fonts.heading)
  const body = docFont(kit.fonts.docBody)

  return {
    default: {
      document: { run: { font: body, size: 22, color: ink } },
    },
    paragraphStyles: [
      {
        id: 'DocTitle',
        name: 'Doc Title',
        basedOn: 'Normal',
        next: 'DocBody',
        quickFormat: true,
        run: { font: heading, size: 36, bold: true, color: primary },
        paragraph: { spacing: { before: 0, after: 120 } },
      },
      {
        id: 'DocSubtitle',
        name: 'Doc Subtitle',
        basedOn: 'Normal',
        next: 'DocBody',
        quickFormat: true,
        run: { font: heading, size: 22, color: deep },
        paragraph: {
          spacing: { before: 0, after: 240 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: accent, space: 4 } },
        },
      },
      {
        id: 'DocH1',
        name: 'Doc Heading 1',
        basedOn: 'Normal',
        next: 'DocBody',
        quickFormat: true,
        run: { font: heading, size: 28, bold: true, color: primary },
        paragraph: { spacing: { before: 320, after: 120 }, keepNext: true, outlineLevel: 0 },
      },
      {
        id: 'DocH2',
        name: 'Doc Heading 2',
        basedOn: 'Normal',
        next: 'DocBody',
        quickFormat: true,
        run: { font: heading, size: 24, bold: true, color: deep },
        paragraph: { spacing: { before: 240, after: 80 }, keepNext: true, outlineLevel: 1 },
      },
      {
        id: 'DocH3',
        name: 'Doc Heading 3',
        basedOn: 'Normal',
        next: 'DocBody',
        quickFormat: true,
        run: { font: heading, size: 22, bold: true, color: deep },
        paragraph: { spacing: { before: 160, after: 60 }, keepNext: true, outlineLevel: 2 },
      },
      {
        id: 'DocBody',
        name: 'Doc Body',
        basedOn: 'Normal',
        quickFormat: true,
        run: { font: body, size: 22, color: ink },
        paragraph: { spacing: { after: 120, line: 276 } },
      },
      {
        id: 'DocSmall',
        name: 'Doc Small',
        basedOn: 'DocBody',
        run: { size: 16, color: mutedInk(kit) },
        paragraph: { spacing: { after: 60 } },
      },
      {
        id: 'DocTableHead',
        name: 'Doc Table Head',
        basedOn: 'DocBody',
        run: { size: 18, bold: true, color: 'FFFFFF' },
        paragraph: { spacing: { before: 40, after: 40 } },
      },
      {
        id: 'DocTableCell',
        name: 'Doc Table Cell',
        basedOn: 'DocBody',
        run: { size: 18 },
        paragraph: { spacing: { before: 40, after: 40 } },
      },
      {
        id: 'DocFooter',
        name: 'Doc Footer',
        basedOn: 'DocBody',
        run: { size: 16, color: deep },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 } },
      },
    ],
  }
}
