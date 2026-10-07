/**
 * docx.layout — report-level layout blocks built on the primitives: headline
 * stat tiles, an internal-link contents list, and a running page header.
 * (Cover pages stay in the consuming app: they carry app-specific content —
 * site photo, customer lockup, tier blurb — and compose from these.)
 */
import { Header, Paragraph, Table } from 'docx';
import type { TenantBrandKit } from '@eq-solutions/contracts';
import { type StatusKind } from './primitives.js';
export interface KpiTile {
    /** Small caption under the figure, e.g. "PASS RATE". */
    label: string;
    /** The headline figure, e.g. "100%" or "04". */
    value: string;
    /** Optional detail line, e.g. "4 / 4 tasks". */
    sub?: string;
    /** Colours the figure with the fixed result colour; omit for the kit primary. */
    status?: StatusKind;
}
/** Row of headline stat tiles (big figure, caption, optional detail) — a complete Table. */
export declare function kpiRow(kit: TenantBrandKit, tiles: KpiTile[]): Table;
export interface TocEntry {
    label: string;
    /** Bookmark name set via `h1(text, { bookmark })` / `h2`. */
    anchor: string;
    /** Indent as a sub-entry. */
    indent?: boolean;
}
/** Contents list: a heading plus one clickable internal link per entry. Static (no Word field to refresh). */
export declare function toc(kit: TenantBrandKit, entries: TocEntry[], title?: string): Paragraph[];
export interface PageHeaderOptions {
    left: string;
    /** Right-aligned text, e.g. a report title or period. */
    right?: string;
}
/** Running page header: small left text, optional right text, rule below in the kit primary. */
export declare function pageHeader(kit: TenantBrandKit, opts: PageHeaderOptions): Header;
