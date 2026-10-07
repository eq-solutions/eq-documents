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
/**
 * Fixed regulatory-alert amber, independent of tenant palette. Deliberately
 * NOT kit-driven: an alert (e.g. a WHS Reg Schedule 3 High-Risk Construction
 * Work flag) needs to look the same for every tenant and never blend into
 * that tenant's own branded fields — the exact reason eq-field hand-coded
 * this colour in v3.5.576 before this kit existed. Same treatment as HAIRLINE.
 */
export declare const ALERT_AMBER = "D97706";
/**
 * Fixed pass / fail / warn colours, independent of tenant palette — same
 * reasoning as ALERT_AMBER: a result must read the same for every tenant.
 * TINT = light cell fill (the table's own ink text stays legible on it);
 * COLOR = solid, for large figures (kpiRow values).
 */
export declare const STATUS_TINT: {
    readonly pass: "DCFCE7";
    readonly fail: "FEE2E2";
    readonly warn: "FEF3C7";
};
export declare const STATUS_COLOR: {
    readonly pass: "16A34A";
    readonly fail: "DC2626";
    readonly warn: "D97706";
};
export type StatusKind = keyof typeof STATUS_TINT;
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
    /** Display width of the logo in px. Default 180. Clamped to the logo column's own budget — see MASTHEAD_LOGO_COLUMN_PX. */
    logoWidthPx?: number;
}
/**
 * Masthead: title + subtitle on the left, logo on the right, in a borderless
 * two-column table. Uses DocTitle / DocSubtitle styles.
 */
export declare function masthead(kit: TenantBrandKit, opts: MastheadOptions): Table;
/** Footer line from kit.legal: `Legal Name | ABN … | address | phone`. Omits missing parts, never invents them. */
export declare function footerText(kit: TenantBrandKit): string;
export interface FooterOptions {
    /** Append ` | Page {current} of {total}` using docx's native PAGE/NUMPAGES fields. Default false — not every document wants pagination (e.g. a one-page confirmation). */
    pageNumbers?: boolean;
}
/** Page footer with the legal line in DocFooter style, a primary-colour top rule, and optional page numbering. */
export declare function footer(kit: TenantBrandKit, opts?: FooterOptions): Footer;
export interface HeadingOptions {
    /** Start the heading on a new page (e.g. one section per asset). */
    pageBreakBefore?: boolean;
    /** Bookmark name so `toc()` (or any internal link) can jump here. */
    bookmark?: string;
}
export declare function h1(text: string, opts?: HeadingOptions): Paragraph;
export declare function h2(text: string, opts?: HeadingOptions): Paragraph;
export declare function body(text: string): Paragraph;
export declare function small(text: string): Paragraph;
export declare function spacer(): Paragraph;
/**
 * Header fill colour for dataTable()'s header row: the kit's primary, unless
 * its luminance falls outside a legible "fill" band — then falls back to
 * palette.deep, which every kit already carries and which the rest of the
 * table family (footer's top rule, DocH2, DocSubtitle) already treats as a
 * legitimate brand colour.
 */
export declare function tableHeadFill(kit: TenantBrandKit): string;
/**
 * A body cell: plain text, `{ text, bold }` to emphasise it (e.g. a failed audit
 * answer), or `{ text, status }` to tint a result cell pass / fail / warn. Colour
 * never comes from the caller: `status` maps to the fixed STATUS_TINT set (the
 * same exception class as ALERT_AMBER, allow-listed in preflight), so a tenant's
 * brand and the palette check are never bypassed by an arbitrary per-cell colour.
 * `{ text, progress: { done, total } }` prefixes the text with a 20-segment bar
 * (e.g. text '3/4' for a 3-of-4 completion cell).
 */
export type DataTableCell = string | {
    text: string;
    bold?: boolean;
    status?: StatusKind;
    progress?: {
        done: number;
        total: number;
    };
};
export interface DataTableOptions {
    head: string[];
    rows: DataTableCell[][];
    /** Column widths in percent; defaults to equal. */
    widths?: number[];
    /** Zebra-stripe body rows with kit.palette.ice. Default true. */
    zebra?: boolean;
}
/** Header row in primary fill with white text, zebra body rows in ice, hairline borders. The "white on blue" rule both briefs share. */
export declare function dataTable(kit: TenantBrandKit, opts: DataTableOptions): Table;
export interface AlertTableOptions {
    /** One full-width amber row per line — e.g. one row per selected HRCW category. */
    lines: string[];
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
export declare function alertTable(opts: AlertTableOptions): Table;
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
export interface SignatureAttendee {
    name: string;
    sig?: {
        bytes: Uint8Array | ArrayBuffer | Buffer;
        mime: ImageMime;
    };
}
export interface SignatureGridOptions {
    attendees: SignatureAttendee[];
    /** Columns per row. Default 2 (the crew sign-off grid both current generators use). */
    columns?: number;
    sigWidthPx?: number;
    sigHeightPx?: number;
}
/**
 * Signature/attendance grid, N-up (default 2) — a complete Table, so a
 * no-bundler consumer (eq-field) that has no access to the raw docx classes
 * (Table/TableRow) can still build a crew sign-off sheet from signatureCell
 * without them. Unlike signatureCell (a bare TableCell for a consumer that
 * already has Table/TableRow), this is the primitive Field actually calls.
 */
export declare function signatureGrid(kit: TenantBrandKit, opts: SignatureGridOptions): Table;
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
