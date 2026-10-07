/**
 * docx.primitives — the shared vocabulary every base document is built from.
 * eq-field's docx-builder.js (docTitle, brandedHeading, secHead, tblOpen*,
 * declaration, imgRun…) and eq-service's lib/reports/* (masthead, kv tables,
 * checklists) both already have versions of these; this is the one copy.
 *
 * Every primitive takes the kit explicitly. Nothing here reads globals.
 */
import {
  AlignmentType,
  BorderStyle,
  Bookmark,
  Footer,
  ImageRun,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  type IBorderOptions,
  type ISectionOptions,
} from 'docx'
import type { LogoAsset, TenantBrandKit } from '@eq-solutions/contracts'
import { luminance, textOn } from '../brand/contrast.js'
import { softFill } from './styles.js'

/** Neutral hairline used for table borders in both brand briefs. Not a brand colour. */
export const HAIRLINE = 'CCCCCC'

/**
 * Fixed regulatory-alert amber, independent of tenant palette. Deliberately
 * NOT kit-driven: an alert (e.g. a WHS Reg Schedule 3 High-Risk Construction
 * Work flag) needs to look the same for every tenant and never blend into
 * that tenant's own branded fields — the exact reason eq-field hand-coded
 * this colour in v3.5.576 before this kit existed. Same treatment as HAIRLINE.
 */
export const ALERT_AMBER = 'D97706'

/**
 * Fixed pass / fail / warn colours, independent of tenant palette — same
 * reasoning as ALERT_AMBER: a result must read the same for every tenant.
 * TINT = light cell fill (the table's own ink text stays legible on it);
 * COLOR = solid, for large figures (kpiRow values).
 */
export const STATUS_TINT = { pass: 'DCFCE7', fail: 'FEE2E2', warn: 'FEF3C7' } as const
export const STATUS_COLOR = { pass: '16A34A', fail: 'DC2626', warn: ALERT_AMBER } as const
export type StatusKind = keyof typeof STATUS_TINT

const hairline: IBorderOptions = { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE }
const borders = { top: hairline, bottom: hairline, left: hairline, right: hairline }
const noBorder: IBorderOptions = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder }

/** A4 portrait with 2 cm margins. Pass as `properties` on a section. */
export function pageA4(): NonNullable<ISectionOptions['properties']> {
  return {
    page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
    },
  }
}

/** docPr name/descr/title stamped on every logoRun so preflight's extractFacts can tell a kit
 *  logo apart from a signature or photo image and grade only the logo's aspect ratio. */
const KIT_LOGO_ALT_TEXT = { title: 'kit-logo', description: 'kit-logo', name: 'kit-logo' }

/** Emit a logo run at a given display width, deriving height from the STORED aspect ratio. Never re-measures. */
export function logoRun(asset: LogoAsset, bytes: Uint8Array | ArrayBuffer | Buffer, widthPx: number): ImageRun {
  const heightPx = Math.round(widthPx / (asset.widthPx / asset.heightPx))
  const type = asset.mime === 'image/svg+xml' ? 'svg' : asset.mime === 'image/jpeg' ? 'jpg' : 'png'
  if (type === 'svg') {
    // docx requires a raster fallback for SVG; consumers should rasterise at upload. Until then, treat as png bytes.
    return new ImageRun({ type: 'png', data: bytes as Buffer, transformation: { width: widthPx, height: heightPx }, altText: KIT_LOGO_ALT_TEXT })
  }
  return new ImageRun({ type, data: bytes as Buffer, transformation: { width: widthPx, height: heightPx }, altText: KIT_LOGO_ALT_TEXT })
}

export type ImageMime = 'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp'

function imageType(mime: ImageMime): 'png' | 'jpg' | 'gif' | 'bmp' {
  return mime === 'image/jpeg' ? 'jpg' : (mime.slice('image/'.length) as 'png' | 'gif' | 'bmp')
}

export interface ImageRunOptions {
  bytes: Uint8Array | ArrayBuffer | Buffer
  mime: ImageMime
  widthPx: number
  heightPx: number
}

/**
 * Generic inline image run at a caller-given display size — unlike logoRun,
 * takes no LogoAsset and does no aspect-ratio derivation; the caller decides
 * the box (matches docx-builder.js's imgRun, which places signatures/photos
 * at a fixed EMU size regardless of the source image's own dimensions).
 * Deliberately untagged (no altText) so preflight's logo-ratio check ignores it.
 */
export function imageRun(opts: ImageRunOptions): ImageRun {
  return new ImageRun({ type: imageType(opts.mime), data: opts.bytes as Buffer, transformation: { width: opts.widthPx, height: opts.heightPx } })
}

export interface MastheadOptions {
  title: string
  subtitle?: string
  /** Bytes of kit.logos.light, fetched by the caller. Omit for no logo (neutral kits). */
  logoBytes?: Uint8Array | ArrayBuffer | Buffer
  /** Display width of the logo in px. Default 180. Clamped to the logo column's own budget — see MASTHEAD_LOGO_COLUMN_PX. */
  logoWidthPx?: number
}

/**
 * A4 content width (pageA4()'s 11906-twip page minus its 1134-twip margins
 * each side), in px at 96 dpi (1440 twips/in ÷ 96 px/in = 15 twips/px).
 * Duplicated as a constant rather than reading pageA4() because its return
 * type allows string page-measure units that don't support arithmetic —
 * keep this in sync if pageA4()'s page size or margins ever change.
 */
const A4_CONTENT_WIDTH_TWIPS = 11906 - 2 * 1134
const A4_CONTENT_WIDTH_PX = A4_CONTENT_WIDTH_TWIPS / 15

/**
 * Available width of masthead()'s 32%-wide logo cell, in px, with a 10%
 * allowance for the cell's own default padding. docx doesn't shrink an
 * ImageRun's fixed EMU size to fit its cell — a requested logo width bigger
 * than this budget overflows into (or clips against) the title column
 * instead of erroring, so masthead() clamps to it below.
 */
const MASTHEAD_LOGO_COLUMN_PX = A4_CONTENT_WIDTH_PX * 0.32 * 0.9

/**
 * Masthead: title + subtitle on the left, logo on the right, in a borderless
 * two-column table. Uses DocTitle / DocSubtitle styles.
 */
export function masthead(kit: TenantBrandKit, opts: MastheadOptions): Table {
  const left: Paragraph[] = [new Paragraph({ style: 'DocTitle', children: [new TextRun(opts.title)] })]
  if (opts.subtitle) left.push(new Paragraph({ style: 'DocSubtitle', children: [new TextRun(opts.subtitle)] }))
  const right: Paragraph[] = []
  if (opts.logoBytes && kit.logos.light) {
    const widthPx = Math.min(opts.logoWidthPx ?? 180, MASTHEAD_LOGO_COLUMN_PX)
    right.push(
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [logoRun(kit.logos.light, opts.logoBytes, widthPx)],
      }),
    )
  }
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({ width: { size: 68, type: WidthType.PERCENTAGE }, borders: noBorders, verticalAlign: VerticalAlign.CENTER, children: left }),
          new TableCell({ width: { size: 32, type: WidthType.PERCENTAGE }, borders: noBorders, verticalAlign: VerticalAlign.CENTER, children: right.length ? right : [new Paragraph('')] }),
        ],
      }),
    ],
  })
}

/** Footer line from kit.legal: `Legal Name | ABN … | address | phone`. Omits missing parts, never invents them. */
export function footerText(kit: TenantBrandKit): string {
  const parts = [kit.tenant.legalName]
  if (kit.legal.abn) parts.push(`ABN ${kit.legal.abn}`)
  if (kit.legal.address) parts.push(kit.legal.address)
  if (kit.legal.phone) parts.push(kit.legal.phone)
  if (kit.legal.web) parts.push(kit.legal.web)
  return parts.join('  |  ')
}

export interface FooterOptions {
  /** Append ` | Page {current} of {total}` using docx's native PAGE/NUMPAGES fields. Default false — not every document wants pagination (e.g. a one-page confirmation). */
  pageNumbers?: boolean
}

/** Page footer with the legal line in DocFooter style, a primary-colour top rule, and optional page numbering. */
export function footer(kit: TenantBrandKit, opts: FooterOptions = {}): Footer {
  const children = [new TextRun(footerText(kit))]
  if (opts.pageNumbers) {
    children.push(new TextRun('  |  Page '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' of '), new TextRun({ children: [PageNumber.TOTAL_PAGES] }))
  }
  return new Footer({
    children: [
      new Paragraph({
        style: 'DocFooter',
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: kit.palette.primary, space: 4 } },
        children,
      }),
    ],
  })
}

export interface HeadingOptions {
  /** Start the heading on a new page (e.g. one section per asset). */
  pageBreakBefore?: boolean
  /** Bookmark name so `toc()` (or any internal link) can jump here. */
  bookmark?: string
}

function heading(style: 'DocH1' | 'DocH2' | 'DocH3', text: string, opts: HeadingOptions = {}): Paragraph {
  const run = new TextRun(text)
  return new Paragraph({
    style,
    pageBreakBefore: opts.pageBreakBefore,
    children: opts.bookmark ? [new Bookmark({ id: opts.bookmark, children: [run] })] : [run],
  })
}
export function h1(text: string, opts?: HeadingOptions): Paragraph {
  return heading('DocH1', text, opts)
}
export function h2(text: string, opts?: HeadingOptions): Paragraph {
  return heading('DocH2', text, opts)
}
export function h3(text: string, opts?: HeadingOptions): Paragraph {
  return heading('DocH3', text, opts)
}
export function body(text: string): Paragraph {
  return new Paragraph({ style: 'DocBody', children: [new TextRun(text)] })
}
export function small(text: string): Paragraph {
  return new Paragraph({ style: 'DocSmall', children: [new TextRun(text)] })
}
export function spacer(): Paragraph {
  return new Paragraph({ style: 'DocBody', children: [new TextRun('')] })
}

/** Below this, a primary is too pale to read as a filled header band. Above this, it reads as an oversized black bar. */
const HEADER_FILL_LUMINANCE_MIN = 0.05
const HEADER_FILL_LUMINANCE_MAX = 0.75

/**
 * Header fill colour for dataTable()'s header row: the kit's primary, unless
 * its luminance falls outside a legible "fill" band — then falls back to
 * palette.deep, which every kit already carries and which the rest of the
 * table family (footer's top rule, DocH2, DocSubtitle) already treats as a
 * legitimate brand colour.
 */
export function tableHeadFill(kit: TenantBrandKit): string {
  const l = luminance(kit.palette.primary)
  return l < HEADER_FILL_LUMINANCE_MIN || l > HEADER_FILL_LUMINANCE_MAX ? kit.palette.deep : kit.palette.primary
}

function headCell(kit: TenantBrandKit, text: string, widthPct?: number): TableCell {
  const fill = tableHeadFill(kit)
  return new TableCell({
    width: widthPct ? { size: widthPct, type: WidthType.PERCENTAGE } : undefined,
    borders,
    shading: { type: ShadingType.CLEAR, fill, color: 'auto' },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ style: 'DocTableHead', keepNext: true, children: [new TextRun({ text, color: textOn(fill, kit.palette.ink) })] })],
  })
}

function bodyCell(kit: TenantBrandKit, cell: DataTableCell, zebra: boolean, widthPct?: number, keepNext = false): TableCell {
  const text = typeof cell === 'string' ? cell : cell.text
  const status = typeof cell === 'string' ? undefined : cell.status
  // A status cell is always bold: the tint alone must never carry the result (print, colour-blind readers).
  const bold = status !== undefined || (typeof cell !== 'string' && cell.bold === true)
  const fill = status ? STATUS_TINT[status] : zebra ? softFill(kit) : undefined
  const progress = typeof cell === 'string' ? undefined : cell.progress
  const runs = progress
    ? [...progressBarRuns(kit, progress), new TextRun({ text: `  ${text}`, bold: bold || undefined })]
    : [new TextRun(bold ? { text, bold: true } : text)]
  return new TableCell({
    width: widthPct ? { size: widthPct, type: WidthType.PERCENTAGE } : undefined,
    borders,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ style: 'DocTableCell', keepNext, children: runs })],
  })
}

/** Number of segments in a progress bar. */
const PROGRESS_SEGMENTS = 20

/**
 * Progress bar as shaded runs of non-breaking spaces (no images, no nested
 * tables): filled segments in the table-header fill, the rest in the neutral hairline grey (the track). Both are
 * kit/neutral colours, so preflight needs no exception.
 */
function progressBarRuns(kit: TenantBrandKit, p: { done: number; total: number }): TextRun[] {
  const ratio = p.total > 0 ? Math.min(Math.max(p.done / p.total, 0), 1) : 0
  const filled = Math.round(ratio * PROGRESS_SEGMENTS)
  const seg = (n: number, fill: string) =>
    new TextRun({ text: ' '.repeat(n), shading: { type: ShadingType.CLEAR, fill, color: 'auto' } })
  const runs: TextRun[] = []
  if (filled > 0) runs.push(seg(filled, tableHeadFill(kit)))
  if (filled < PROGRESS_SEGMENTS) runs.push(seg(PROGRESS_SEGMENTS - filled, HAIRLINE))
  return runs
}

/**
 * A body cell: plain text, `{ text, bold }` to emphasise it (e.g. a failed audit
 * answer), or `{ text, status }` to tint a result cell pass / fail / warn. Colour
 * never comes from the caller: `status` maps to the fixed STATUS_TINT set (the
 * same exception class as ALERT_AMBER, allow-listed in preflight), so a tenant's
 * brand and the palette check are never bypassed by an arbitrary per-cell colour.
 * `{ text, progress: { done, total } }` prefixes the text with a 20-segment bar
 * (e.g. text '3/4' for a 3-of-4 completion cell).
 */
export type DataTableCell = string | { text: string; bold?: boolean; status?: StatusKind; progress?: { done: number; total: number } }

/** Body row `ri` of `n` keeps with the row after it: the first two, and the second-to-last. */
function keepWithNext(ri: number, n: number): boolean {
  return (ri < 2 && ri < n - 1) || (n >= 3 && ri === n - 2)
}

export interface DataTableOptions {
  head: string[]
  rows: DataTableCell[][]
  /** Column widths in percent; defaults to equal. */
  widths?: number[]
  /** Zebra-stripe body rows with kit.palette.ice. Default true. */
  zebra?: boolean
}

/** Header row in primary fill with white text, zebra body rows in ice, hairline borders. The "white on blue" rule both briefs share. */
export function dataTable(kit: TenantBrandKit, opts: DataTableOptions): Table {
  const n = opts.head.length
  const widths = opts.widths ?? opts.head.map(() => Math.floor(100 / n))
  const zebra = opts.zebra ?? true
  // Fixed layout + explicit grid: with only per-cell percentage widths Word (and
  // LibreOffice) autofit columns to their content, so a 4-column checklist
  // rendered lopsided and a narrow table stopped short of the margin.
  const widthTotal = widths.reduce((sum, w) => sum + w, 0) || 100 // callers' widths needn't sum to exactly 100
  const columnWidths = widths.map((w) => Math.round((A4_CONTENT_WIDTH_TWIPS * w) / widthTotal))
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: opts.head.map((h, i) => headCell(kit, h, widths[i])) }),
      // Header + first two body rows stay with the row after them (a table starting at the foot of a
      // page moves whole instead of stranding its header and one row); the second-to-last row too, so
      // the last row is never alone on the next page.
      ...opts.rows.map(
        (r, ri) =>
          new TableRow({ cantSplit: true, children: r.map((c, ci) => bodyCell(kit, c, zebra && ri % 2 === 1, widths[ci], keepWithNext(ri, opts.rows.length))) }),
      ),
    ],
  })
}

export interface AlertTableOptions {
  /** One full-width amber row per line — e.g. one row per selected HRCW category. */
  lines: string[]
}

function alertCell(text: string): TableCell {
  return new TableCell({
    borders,
    shading: { type: ShadingType.CLEAR, fill: ALERT_AMBER, color: 'auto' },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ style: 'DocTableCell', children: [new TextRun({ text, bold: true, color: 'FFFFFF' })] })],
  })
}

/**
 * Alert table: one full-width amber row per line, bold white text, for a
 * regulatory/safety flag that must stand out regardless of tenant brand —
 * e.g. WHS Reg Schedule 3 High-Risk Construction Work categories in a
 * Prestart. Not part of the "white on kit-primary" dataTable/kvTable
 * family: this colour never comes from the kit. Returns a complete Table
 * (like signatureGrid/photoGrid), so a no-bundler consumer (eq-field) with
 * no access to the raw docx Table/TableRow classes can drop it straight
 * into a section's children alongside dataTable/kvTable, rather than
 * needing to merge rows into another table.
 */
export function alertTable(opts: AlertTableOptions): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: opts.lines.map((text) => new TableRow({ cantSplit: true, children: [alertCell(text)] })),
  })
}

/** Two-column label/value table: label cells in ice, values plain. The "kvTable" both Field and Service draw by hand today. */
export function kvTable(kit: TenantBrandKit, pairs: Array<[string, string]>, labelWidthPct = 30): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: pairs.map(
      ([k, v]) =>
        new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: labelWidthPct, type: WidthType.PERCENTAGE },
              borders,
              shading: { type: ShadingType.CLEAR, fill: softFill(kit), color: 'auto' },
              children: [new Paragraph({ style: 'DocTableCell', children: [new TextRun({ text: k, bold: true })] })],
            }),
            new TableCell({
              width: { size: 100 - labelWidthPct, type: WidthType.PERCENTAGE },
              borders,
              children: [new Paragraph({ style: 'DocTableCell', children: [new TextRun(v)] })],
            }),
          ],
        }),
    ),
  })
}

export interface SignatureCellOptions {
  name: string
  /** Signature image bytes + mime. Omit for an unsigned attendee (blank line, same cell height). */
  sig?: { bytes: Uint8Array | ArrayBuffer | Buffer; mime: ImageMime }
  /** Display box for the signature image. Defaults match docx-builder.js's imgRun default (~1.5"×0.5" @96dpi). */
  sigWidthPx?: number
  sigHeightPx?: number
  widthPct?: number
}

/** Signature/attendance cell: bold name, then the signature image or a blank line. Ice-shaded, same "white on blue" table family as dataTable/kvTable. */
export function signatureCell(kit: TenantBrandKit, opts: SignatureCellOptions): TableCell {
  const nameP = new Paragraph({ style: 'DocTableCell', children: [new TextRun({ text: opts.name, bold: true })] })
  const sigP = opts.sig
    ? new Paragraph({
        children: [
          imageRun({
            bytes: opts.sig.bytes,
            mime: opts.sig.mime,
            widthPx: opts.sigWidthPx ?? 144,
            heightPx: opts.sigHeightPx ?? 48,
          }),
        ],
      })
    : new Paragraph({ style: 'DocTableCell', children: [new TextRun('')] })
  return new TableCell({
    width: opts.widthPct ? { size: opts.widthPct, type: WidthType.PERCENTAGE } : undefined,
    borders,
    shading: { type: ShadingType.CLEAR, fill: softFill(kit), color: 'auto' },
    verticalAlign: VerticalAlign.CENTER,
    children: [nameP, sigP],
  })
}

export interface SignatureAttendee {
  name: string
  sig?: { bytes: Uint8Array | ArrayBuffer | Buffer; mime: ImageMime }
}

export interface SignatureGridOptions {
  attendees: SignatureAttendee[]
  /** Columns per row. Default 2 (the crew sign-off grid both current generators use). */
  columns?: number
  sigWidthPx?: number
  sigHeightPx?: number
}

function blankIceCell(kit: TenantBrandKit, widthPct: number): TableCell {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    borders,
    shading: { type: ShadingType.CLEAR, fill: softFill(kit), color: 'auto' },
    children: [new Paragraph('')],
  })
}

/**
 * Signature/attendance grid, N-up (default 2) — a complete Table, so a
 * no-bundler consumer (eq-field) that has no access to the raw docx classes
 * (Table/TableRow) can still build a crew sign-off sheet from signatureCell
 * without them. Unlike signatureCell (a bare TableCell for a consumer that
 * already has Table/TableRow), this is the primitive Field actually calls.
 */
export function signatureGrid(kit: TenantBrandKit, opts: SignatureGridOptions): Table {
  const cols = opts.columns ?? 2
  const widthPct = Math.floor(100 / cols)
  const rows: TableRow[] = []
  for (let i = 0; i < opts.attendees.length; i += cols) {
    const rowAttendees = opts.attendees.slice(i, i + cols)
    const cells: TableCell[] = []
    for (let c = 0; c < cols; c++) {
      const a = rowAttendees[c]
      cells.push(
        a
          ? signatureCell(kit, { name: a.name, sig: a.sig, sigWidthPx: opts.sigWidthPx, sigHeightPx: opts.sigHeightPx, widthPct })
          : blankIceCell(kit, widthPct),
      )
    }
    rows.push(new TableRow({ cantSplit: true, children: cells }))
  }
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows })
}

export interface PhotoGridPhoto {
  bytes: Uint8Array | ArrayBuffer | Buffer
  mime: ImageMime
  caption?: string
}

export interface PhotoGridOptions {
  photos: PhotoGridPhoto[]
  /** Columns per row. Default 2 (the "2-up" grid both current generators use). */
  columns?: number
  /** Display box per photo. Defaults match docx-builder.js's imgRun default for photos (3"×2.25" @96dpi). */
  photoWidthPx?: number
  photoHeightPx?: number
}

function photoCell(photo: PhotoGridPhoto | undefined, widthPx: number, heightPx: number, widthPct: number): TableCell {
  if (!photo) {
    return new TableCell({ width: { size: widthPct, type: WidthType.PERCENTAGE }, borders: noBorders, children: [new Paragraph('')] })
  }
  const children: Paragraph[] = [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [imageRun({ bytes: photo.bytes, mime: photo.mime, widthPx, heightPx })] }),
  ]
  if (photo.caption) children.push(new Paragraph({ style: 'DocSmall', alignment: AlignmentType.CENTER, children: [new TextRun(photo.caption)] }))
  return new TableCell({ width: { size: widthPct, type: WidthType.PERCENTAGE }, borders: noBorders, children })
}

/** Photo grid, N-up (default 2), borderless cells, each photo centred with an optional caption below it. The "2-up photo grid" both Field and Service already hand-build. */
export function photoGrid(kit: TenantBrandKit, opts: PhotoGridOptions): Table {
  const cols = opts.columns ?? 2
  const widthPct = Math.floor(100 / cols)
  const widthPx = opts.photoWidthPx ?? 288
  const heightPx = opts.photoHeightPx ?? 216
  const rows: TableRow[] = []
  for (let i = 0; i < opts.photos.length; i += cols) {
    const rowPhotos = opts.photos.slice(i, i + cols)
    const cells: TableCell[] = []
    for (let c = 0; c < cols; c++) cells.push(photoCell(rowPhotos[c], widthPx, heightPx, widthPct))
    rows.push(new TableRow({ cantSplit: true, children: cells }))
  }
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: noBorders, rows })
}
