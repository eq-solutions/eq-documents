import type { TenantBrandKit } from '@eq-solutions/contracts';
export type CheckId = 'logo' | 'ratio' | 'palette' | 'fonts' | 'flat' | 'footer';
export interface CheckResult {
    id: CheckId;
    pass: boolean;
    detail?: string;
}
export interface PreflightResult {
    ok: boolean;
    checks: CheckResult[];
    /** The one-line chat form: "Brand check: ✓ logo ✓ ratio ✗ palette → …" */
    line: string;
}
/** What a document declares about itself. Fill from the generator, or let preflightDocx extract it. */
export interface DocumentFacts {
    /** URLs or storage paths of every image placed as a logo. Empty when the kit has no logo. */
    logoSources?: string[];
    /** Placed logo sizes (px or EMU — only the ratio matters). */
    logoPlacements?: Array<{
        width: number;
        height: number;
    }>;
    /** Every colour hex used for text, fills, borders. Uppercase bare hex. */
    usedHex: string[];
    /** Every font family name used. */
    usedFonts: string[];
    /** Any gradient / shadow / glow / effect present. */
    hasEffects: boolean;
    /** Footer part present and containing the legal name. */
    hasFooter: boolean;
    footerText?: string;
}
/** Colours a kit legitimately permits beyond its palette. */
export declare function allowedHex(kit: TenantBrandKit): Set<string>;
export declare function preflight(kit: TenantBrandKit, facts: DocumentFacts): PreflightResult;
/** Extract DocumentFacts from a generated .docx. Logo sources are reported by media part name because the URL is gone once embedded. */
export declare function extractFacts(docx: Uint8Array | ArrayBuffer | Blob): Promise<DocumentFacts>;
/** Run the six checks against the bytes of a generated .docx. */
export declare function preflightDocx(kit: TenantBrandKit, docx: Uint8Array | ArrayBuffer | Blob): Promise<PreflightResult>;
