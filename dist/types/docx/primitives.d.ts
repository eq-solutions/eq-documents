/**
 * docx.primitives — the shared vocabulary every base document is built from.
 * eq-field's docx-builder.js (docTitle, brandedHeading, secHead, tblOpen*,
 * declaration, imgRun…) and eq-service's lib/reports/* (masthead, kv tables,
 * checklists) both already have versions of these; this is the one copy.
 *
 * Every primitive takes the kit explicitly. Nothing here reads globals.
 */
import { Footer, ImageRun, Paragraph, Table, TableCell, type ISectionOptions } from 'docx';
import type { LogoAsset, TenantBrandKit } from '@eq-solutions/contracts';
/** Neutral hairline used for table borders in both brand briefs. Not a brand colour. */
export declare const HAIRLINE = "CCCCCC";
/** A4 portrait with 2 cm margins. Pass as `properties` on a section. */
export declare function pageA4(): NonNullable<ISectionOptions['properties']>;
/** Emit a logo run at a given display width, deriving height from the STORED aspect ratio. Never re-measures. */
export declare function logoRun(asset: LogoAsset, bytes: Uint8Array | ArrayBuffer | Buffer, widthPx: number): ImageRun;
export type ImageMime = 'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp';
export interface ImageRunOptions {
    bytes: Uint8Array | ArrayBuffer | Buffer;
    mime: ImageMime;
    widthPx: number;
    heightPx: number;
}
/**
 * Generic inline image run at a caller-given display size — unlike logoRun,
 * takes no LogoAsset and does no aspect-ratio derivation; the caller decides
 * the box (matches docx-builder.js's imgRun, which places signatures/photos
 * at a fixed EMU size regardless of the source image's own dimensions).
 * Deliberately untagged (no altText) so preflight's logo-ratio check ignores it.
 */
export declare function imageRun(opts: ImageRunOptions): ImageRun;
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
export interface SignatureCellOptions {
    name: string;
    /** Signature image bytes + mime. Omit for an unsigned attendee (blank line, same cell height). */
    sig?: {
        bytes: Uint8Array | ArrayBuffer | Buffer;
        mime: ImageMime;
    };
    /** Display box for the signature image. Defaults match docx-builder.js's imgRun default (~1.5"×0.5" @96dpi). */
    sigWidthPx?: number;
    sigHeightPx?: number;
    widthPct?: number;
}
/** Signature/attendance cell: bold name, then the signature image or a blank line. Ice-shaded, same "white on blue" table family as dataTable/kvTable. */
export declare function signatureCell(kit: TenantBrandKit, opts: SignatureCellOptions): TableCell;
export interface PhotoGridPhoto {
    bytes: Uint8Array | ArrayBuffer | Buffer;
    mime: ImageMime;
    caption?: string;
}
export interface PhotoGridOptions {
    photos: PhotoGridPhoto[];
    /** Columns per row. Default 2 (the "2-up" grid both current generators use). */
    columns?: number;
    /** Display box per photo. Defaults match docx-builder.js's imgRun default for photos (3"×2.25" @96dpi). */
    photoWidthPx?: number;
    photoHeightPx?: number;
}
/** Photo grid, N-up (default 2), borderless cells, each photo centred with an optional caption below it. The "2-up photo grid" both Field and Service already hand-build. */
export declare function photoGrid(kit: TenantBrandKit, opts: PhotoGridOptions): Table;
