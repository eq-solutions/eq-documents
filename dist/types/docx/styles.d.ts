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
import type { IStylesOptions } from 'docx';
import type { TenantBrandKit } from '@eq-solutions/contracts';
/** Word has no font-family fallback list — `w:rFonts` names exactly one face per script range (ascii/hAnsi/eastAsia/cs), so an unrecognised font can't degrade to a chain the way a CSS font-family stack would. */
export declare const FALLBACK_FONT = "Arial";
export declare const DOC_STYLE_IDS: readonly ["DocTitle", "DocSubtitle", "DocH1", "DocH2", "DocH3", "DocBody", "DocSmall", "DocTableHead", "DocTableCell", "DocFooter"];
export type DocStyleId = (typeof DOC_STYLE_IDS)[number];
/** Accent colour: explicit accent role, else deep. */
export declare function accentOf(kit: TenantBrandKit): string;
/** 60 % ink blended towards white — for captions/metadata without inventing a new brand colour. */
export declare function mutedInk(kit: TenantBrandKit): string;
/**
 * Light surface fill for zebra rows, label columns and stat tiles. Normally the
 * kit's `ice`; but a tenant's `ice` can be a mid-tone (SKS's renders as a heavy
 * grey-lavender slab), so when it is too dark to read as a soft surface it is
 * replaced by the primary blended 90 % towards white — the same idea as
 * `tableHeadFill`'s fallback, and still a pure function of the kit.
 */
export declare function softFill(kit: TenantBrandKit): string;
/** Build the `styles` option for `new Document({ styles })`. */
export declare function docxStyles(kit: TenantBrandKit): IStylesOptions;
