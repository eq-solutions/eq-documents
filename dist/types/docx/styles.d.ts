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
export declare const DOC_STYLE_IDS: readonly ["DocTitle", "DocSubtitle", "DocH1", "DocH2", "DocBody", "DocSmall", "DocTableHead", "DocTableCell", "DocFooter"];
export type DocStyleId = (typeof DOC_STYLE_IDS)[number];
/** Accent colour: explicit accent role, else deep. */
export declare function accentOf(kit: TenantBrandKit): string;
/** 60 % ink blended towards white — for captions/metadata without inventing a new brand colour. */
export declare function mutedInk(kit: TenantBrandKit): string;
/** Build the `styles` option for `new Document({ styles })`. */
export declare function docxStyles(kit: TenantBrandKit): IStylesOptions;
