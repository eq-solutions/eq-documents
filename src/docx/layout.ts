/**
 * docx.layout — report-level layout blocks built on the primitives: headline
 * stat tiles, an internal-link contents list, and a running page header.
 * (Cover pages stay in the consuming app: they carry app-specific content —
 * site photo, customer lockup, tier blurb — and compose from these.)
 */
import {
  AlignmentType,
  BorderStyle,
  Header,
  InternalHyperlink,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TabStopType,
  TextRun,
  VerticalAlign,
  WidthType,
  type IBorderOptions,
} from 'docx'
import type { TenantBrandKit } from '@eq-solutions/contracts'
import { HAIRLINE, STATUS_COLOR, h1, type StatusKind } from './primitives.js'
import { mutedInk, softFill } from './styles.js'

const hairline: IBorderOptions = { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE }
const borders = { top: hairline, bottom: hairline, left: hairline, right: hairline }

export interface KpiTile {
  /** Small caption under the figure, e.g. "PASS RATE". */
  label: string
  /** The headline figure, e.g. "100%" or "04". */
  value: string
  /** Optional detail line, e.g. "4 / 4 tasks". */
  sub?: string
  /** Colours the figure with the fixed result colour; omit for the kit primary. */
  status?: StatusKind
}

/** Row of headline stat tiles (big figure, caption, optional detail) — a complete Table. */
export function kpiRow(kit: TenantBrandKit, tiles: KpiTile[]): Table {
  const widthPct = Math.floor(100 / Math.max(tiles.length, 1))
  const muted = mutedInk(kit)
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        cantSplit: true,
        children: tiles.map(
          (t) =>
            new TableCell({
              width: { size: widthPct, type: WidthType.PERCENTAGE },
              borders,
              shading: { type: ShadingType.CLEAR, fill: softFill(kit), color: 'auto' },
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 120 },
                  children: [new TextRun({ text: t.value, bold: true, size: 44, color: t.status ? STATUS_COLOR[t.status] : kit.palette.primary })],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: t.label.toUpperCase(), bold: true, size: 16, color: muted })],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 120 },
                  children: [new TextRun({ text: t.sub ?? '', size: 16, color: muted })],
                }),
              ],
            }),
        ),
      }),
    ],
  })
}

export interface TocEntry {
  label: string
  /** Bookmark name set via `h1(text, { bookmark })` / `h2`. */
  anchor: string
  /** Indent as a sub-entry. */
  indent?: boolean
}

/** Contents list: a heading plus one clickable internal link per entry. Static (no Word field to refresh). */
export function toc(kit: TenantBrandKit, entries: TocEntry[], title = 'Contents'): Paragraph[] {
  return [
    h1(title),
    ...entries.map(
      (e) =>
        new Paragraph({
          style: 'DocBody',
          indent: e.indent ? { left: 360 } : undefined,
          children: [
            new InternalHyperlink({
              anchor: e.anchor,
              children: [new TextRun({ text: e.label, color: kit.palette.primary, underline: {} })],
            }),
          ],
        }),
    ),
  ]
}

export interface PageHeaderOptions {
  left: string
  /** Right-aligned text, e.g. a report title or period. */
  right?: string
}

/** Running page header: small left text, optional right text, rule below in the kit primary. */
export function pageHeader(kit: TenantBrandKit, opts: PageHeaderOptions): Header {
  const children = [new TextRun(opts.left)]
  if (opts.right) children.push(new TextRun({ text: '\t' + opts.right }))
  return new Header({
    children: [
      new Paragraph({
        style: 'DocFooter',
        tabStops: [{ type: TabStopType.RIGHT, position: 9638 }],
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: kit.palette.primary, space: 4 } },
        children,
      }),
    ],
  })
}
