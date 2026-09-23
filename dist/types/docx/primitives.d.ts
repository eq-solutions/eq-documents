/**
 * docx.primitives — the shared vocabulary every base document is built from.
 * eq-field's docx-builder.js (docTitle, brandedHeading, secHead, tblOpen*,
 * declaration, imgRun…) and eq-service's lib/reports/* (masthead, kv tables,
 * checklists) both already have versions of these; this is the one copy.
 *
 * Every primitive takes the kit explicitly. Nothing here reads globals.
 */
import { Footer, ImageRun, Paragraph, Table, type ISectionOptions } from 'docx';
import type { LogoAsset, TenantBrandKit } from '@eq-solutions/contracts';
/** Neutral hairline used for table borders in both brand briefs. Not a brand colour. */
export declare const HAIRLINE = "CCCCCC";
/** A4 portrait with 2 cm margins. Pass as `properties` on a section. */
export declare function pageA4(): NonNullable<ISectionOptions['properties']>;
/** Emit a logo run at a given display width, deriving height from the STORED aspect ratio. Never re-measures. */
export declare function logoRun(asset: LogoAsset, bytes: Uint8Array | ArrayBuffer | Buffer, widthPx: number): ImageRun;
export interface MastheadOptions {
    title: string;
    subtitle?: string;
    /** Bytes of kit.logos.light, fetched by the caller. Omit for no logo (neutral kits). */
    logoBytes?: Uint8Array | ArrayBuffer | Buffer;
    /** Display width of the logo in px. Default 180. */
    logoWidthPx?: number;
}
/**
 * Masthead: title + subtitle on the left, logo on the right, in a borderless
 * two-column table. Uses DocTitle / DocSubtitle styles.
 */
export declare function masthead(kit: TenantBrandKit, opts: MastheadOptions): Table;
/** Footer line from kit.legal: `Legal Name | ABN … | address | phone`. Omits missing parts, never invents them. */
export declare function footerText(kit: TenantBrandKit): string;
/** Page footer with the legal line in DocFooter style and a primary-colour top rule. */
export declare function footer(kit: TenantBrandKit): Footer;
export declare function h1(text: string): Paragraph;
export declare function h2(text: string): Paragraph;
export declare function body(text: string): Paragraph;
export declare function small(text: string): Paragraph;
export declare function spacer(): Paragraph;
export interface DataTableOptions {
    head: string[];
    rows: string[][];
    /** Column widths in percent; defaults to equal. */
    widths?: number[];
    /** Zebra-stripe body rows with kit.palette.ice. Default true. */
    zebra?: boolean;
}
/** Header row in primary fill with white text, zebra body rows in ice, hairline borders. The "white on blue" rule both briefs share. */
export declare function dataTable(kit: TenantBrandKit, opts: DataTableOptions): Table;
/** Two-column label/value table: label cells in ice, values plain. The "kvTable" both Field and Service draw by hand today. */
export declare function kvTable(kit: TenantBrandKit, pairs: Array<[string, string]>, labelWidthPct?: number): Table;
