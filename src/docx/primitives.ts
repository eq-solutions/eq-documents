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
  Footer,
  ImageRun,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  type IBorderOptions,
  type ISectionOptions,
} from 'docx'
import type { LogoAsset, TenantBrandKit } from '@eq-solutions/contracts'
import { textOn } from '../brand/contrast.js'

/** Neutral hairline used for table borders in both brand briefs. Not a brand colour. */
export const HAIRLINE = 'CCCCCC'

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
  /** Display width of the logo in px. Default 180. */
  logoWidthPx?: number
}

/**
 * Masthead: title + subtitle on the left, logo on the right, in a borderless
 * two-column table. Uses DocTitle / DocSubtitle styles.
 */
export function masthead(kit: TenantBrandKit, opts: MastheadOptions): Table {
  const left: Paragraph[] = [new Paragraph({ style: 'DocTitle', children: [new TextRun(opts.title)] })]
  if (opts.subtitle) left.push(new Paragraph({ style: 'DocSubtitle', children: [new TextRun(opts.subtitle)] }))
  const right: Paragraph[] = []
  if (opts.logoBytes && kit.logos.light) {
    right.push(
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [logoRun(kit.logos.light, opts.logoBytes, opts.logoWidthPx ?? 180)],
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

/** Page footer with the legal line in DocFooter style and a primary-colour top rule. */
export function footer(kit: TenantBrandKit): Footer {
  return new Footer({
    children: [
      new Paragraph({
        style: 'DocFooter',
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: kit.palette.primary, space: 4 } },
        children: [new TextRun(footerText(kit))],
      }),
    ],
  })
}

export function h1(text: string): Paragraph {
  return new Paragraph({ style: 'DocH1', children: [new TextRun(text)] })
}
export function h2(text: string): Paragraph {
  return new Paragraph({ style: 'DocH2', children: [new TextRun(text)] })
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

function headCell(kit: TenantBrandKit, text: string, widthPct?: number): TableCell {
  return new TableCell({
    width: widthPct ? { size: widthPct, type: WidthType.PERCENTAGE } : undefined,
    borders,
    shading: { type: ShadingType.CLEAR, fill: kit.palette.primary, color: 'auto' },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ style: 'DocTableHead', children: [new TextRun({ text, color: textOn(kit.palette.primary, kit.palette.ink) })] })],
  })
}

function bodyCell(kit: TenantBrandKit, text: string, zebra: boolean, widthPct?: number): TableCell {
  return new TableCell({
    width: widthPct ? { size: widthPct, type: WidthType.PERCENTAGE } : undefined,
    borders,
    shading: zebra ? { type: ShadingType.CLEAR, fill: kit.palette.ice, color: 'auto' } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ style: 'DocTableCell', children: [new TextRun(text)] })],
  })
}

export interface DataTableOptions {
  head: string[]
  rows: string[][]
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
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: opts.head.map((h, i) => headCell(kit, h, widths[i])) }),
      ...opts.rows.map(
        (r, ri) =>
          new TableRow({ cantSplit: true, children: r.map((c, ci) => bodyCell(kit, c, zebra && ri % 2 === 1, widths[ci])) }),
      ),
    ],
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
              shading: { type: ShadingType.CLEAR, fill: kit.palette.ice, color: 'auto' },
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
    shading: { type: ShadingType.CLEAR, fill: kit.palette.ice, color: 'auto' },
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
    shading: { type: ShadingType.CLEAR, fill: kit.palette.ice, color: 'auto' },
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
