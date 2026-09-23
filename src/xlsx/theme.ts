/**
 * xlsx.theme — exceljs styling from a TenantBrandKit. Lifted from the one
 * genuinely tenant-branded spreadsheet in the suite today (eq-shell
 * netlify/functions/_shared/compliance-register.ts) so eq-service's exports
 * get the same look. ESM-only: not part of the IIFE build.
 */
import type ExcelJS from 'exceljs'
import type { LogoAsset, TenantBrandKit } from '@eq-solutions/contracts'
import { textOn } from '../brand/contrast.js'

type Fill = ExcelJS.Fill
type Font = Partial<ExcelJS.Font>
type Borders = Partial<ExcelJS.Borders>

export interface XlsxTheme {
  headerFill: Fill
  headerFont: Font
  zebraFill: Fill
  bodyFont: Font
  titleFont: Font
  border: Borders
}

const argb = (hex: string) => `FF${hex}`

export function xlsxTheme(kit: TenantBrandKit): XlsxTheme {
  const { primary, ice, ink } = kit.palette
  const thin: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: argb('CCCCCC') } }
  return {
    headerFill: { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(primary) } },
    headerFont: { name: kit.fonts.docBody, bold: true, color: { argb: argb(textOn(primary, ink)) }, size: 10 },
    zebraFill: { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(ice) } },
    bodyFont: { name: kit.fonts.docBody, color: { argb: argb(ink) }, size: 10 },
    titleFont: { name: kit.fonts.heading, bold: true, color: { argb: argb(primary) }, size: 16 },
    border: { top: thin, bottom: thin, left: thin, right: thin },
  }
}

/** Style a header row: primary fill, contrasting bold text, hairline borders. */
export function applyHeaderRow(row: ExcelJS.Row, theme: XlsxTheme): void {
  row.eachCell((cell) => {
    cell.fill = theme.headerFill
    cell.font = theme.headerFont
    cell.border = theme.border
    cell.alignment = { vertical: 'middle' }
  })
}

/** Style body rows `from`..`to` (1-based, inclusive): body font, zebra on even offsets, hairline borders. */
export function applyBodyRows(ws: ExcelJS.Worksheet, theme: XlsxTheme, from: number, to: number, columns: number): void {
  for (let r = from; r <= to; r++) {
    const row = ws.getRow(r)
    const zebra = (r - from) % 2 === 1
    for (let c = 1; c <= columns; c++) {
      const cell = row.getCell(c)
      cell.font = theme.bodyFont
      cell.border = theme.border
      if (zebra) cell.fill = theme.zebraFill
    }
  }
}

export interface XlsxMastheadOptions {
  title: string
  subtitle?: string
  /** Bytes of kit.logos.light. Omit for no logo. */
  logoBytes?: Buffer | ArrayBuffer
  /** Display width in px. Default 160. Height derives from the stored aspect ratio. */
  logoWidthPx?: number
  /** Number of columns the title should span. Default 6. */
  span?: number
}

/**
 * Title (+ subtitle) in rows 1–2 and the logo anchored top-right of the
 * spanned range. Returns the first free row for data. Height is derived
 * from the STORED dimensions — never re-measured.
 */
export function xlsxMasthead(wb: ExcelJS.Workbook, ws: ExcelJS.Worksheet, kit: TenantBrandKit, opts: XlsxMastheadOptions): number {
  const theme = xlsxTheme(kit)
  const span = opts.span ?? 6
  ws.mergeCells(1, 1, 1, Math.max(1, span - 2))
  const title = ws.getCell(1, 1)
  title.value = opts.title
  title.font = theme.titleFont
  ws.getRow(1).height = 30
  let next = 2
  if (opts.subtitle) {
    ws.mergeCells(2, 1, 2, Math.max(1, span - 2))
    const st = ws.getCell(2, 1)
    st.value = opts.subtitle
    st.font = { ...theme.bodyFont, color: { argb: argb(kit.palette.deep) } }
    next = 3
  }
  if (opts.logoBytes && kit.logos.light) {
    const asset: LogoAsset = kit.logos.light
    const width = opts.logoWidthPx ?? 160
    const height = Math.round(width / (asset.widthPx / asset.heightPx))
    const ext = asset.mime === 'image/jpeg' ? 'jpeg' : 'png'
    const id = wb.addImage({ buffer: opts.logoBytes as unknown as ExcelJS.Buffer, extension: ext })
    ws.addImage(id, { tl: { col: span - 2, row: 0 }, ext: { width, height } })
    const rowsNeeded = Math.ceil(height / 20) + 1
    next = Math.max(next, rowsNeeded + 1)
  }
  return next + 1
}
