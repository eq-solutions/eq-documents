/**
 * xlsx.theme — exceljs styling from a TenantBrandKit. Lifted from the one
 * genuinely tenant-branded spreadsheet in the suite today (eq-shell
 * netlify/functions/_shared/compliance-register.ts) so eq-service's exports
 * get the same look. ESM-only: not part of the IIFE build.
 */
import type ExcelJS from 'exceljs';
import type { TenantBrandKit } from '@eq-solutions/contracts';
type Fill = ExcelJS.Fill;
type Font = Partial<ExcelJS.Font>;
type Borders = Partial<ExcelJS.Borders>;
export interface XlsxTheme {
    headerFill: Fill;
    headerFont: Font;
    zebraFill: Fill;
    bodyFont: Font;
    titleFont: Font;
    border: Borders;
}
export declare function xlsxTheme(kit: TenantBrandKit): XlsxTheme;
/** Style a header row: primary fill, contrasting bold text, hairline borders. */
export declare function applyHeaderRow(row: ExcelJS.Row, theme: XlsxTheme): void;
/** Style body rows `from`..`to` (1-based, inclusive): body font, zebra on even offsets, hairline borders. */
export declare function applyBodyRows(ws: ExcelJS.Worksheet, theme: XlsxTheme, from: number, to: number, columns: number): void;
export interface XlsxMastheadOptions {
    title: string;
    subtitle?: string;
    /** Bytes of kit.logos.light. Omit for no logo. */
    logoBytes?: Buffer | ArrayBuffer;
    /** Display width in px. Default 160. Height derives from the stored aspect ratio. */
    logoWidthPx?: number;
    /** Number of columns the title should span. Default 6. */
    span?: number;
}
/**
 * Title (+ subtitle) in rows 1–2 and the logo anchored top-right of the
 * spanned range. Returns the first free row for data. Height is derived
 * from the STORED dimensions — never re-measured.
 */
export declare function xlsxMasthead(wb: ExcelJS.Workbook, ws: ExcelJS.Worksheet, kit: TenantBrandKit, opts: XlsxMastheadOptions): number;
export {};
