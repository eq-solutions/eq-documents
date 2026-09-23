/**
 * brand.normalise — the ONE mapping from a tenant's raw brand record to a
 * TenantBrandKit. The canonical RPC `eq_get_tenant_brand_kit` applies the
 * same rules in SQL; this is the client/server JS copy so a consumer that
 * already holds the raw `organisations.branding` jsonb (eq-field's
 * `__TENANT_BRANDING_CANON__`, eq-service's JWT claims) never needs a second
 * network call, and so the two can be tested against each other.
 *
 * Rules:
 *  - Palette roles are validated individually; a bad or missing role falls
 *    back to NEUTRAL_BRAND_KIT's grey for that role and marks the kit
 *    incomplete. Never to another tenant's value.
 *  - Structured `branding.logos.*` (with measured dimensions) wins. Legacy
 *    keys (`hubLogo` URL, `gateLogo` / `sidebarLogoHtml` HTML strings,
 *    `gateLogoDark` URL) are recognised and their URL extracted, but a URL
 *    without dimensions cannot become a LogoAsset — it is returned as
 *    `legacyLogos` for the caller to measure (see `withMeasuredLogos`).
 *  - Fonts and legal come from `branding.fonts` / `branding.legal` when
 *    present; otherwise neutral (system fonts, empty footer).
 *  - Auth-gate flags (`rememberMeDays`, `coreOnly`, …) are ignored: not brand.
 */
import { type BrandPalette, type LogoAsset, type TenantBrandKit } from '@eq-solutions/contracts';
export interface TenantIdentity {
    id: string;
    slug: string;
    /** Display name. Falls back to slug. */
    name?: string | null;
    /** Legal entity name. Falls back to `name`. */
    legalName?: string | null;
}
/** Raw `public.organisations.branding` jsonb, as stored. Loosely typed on purpose. */
export type RawBranding = Record<string, unknown> | null | undefined;
export interface LegacyLogos {
    light?: string;
    dark?: string;
}
export interface NormaliseResult {
    kit: TenantBrandKit;
    /** Logo URLs found only in legacy keys, without measured dimensions. Measure and merge with withMeasuredLogos(). */
    legacyLogos: LegacyLogos;
    /** Which palette roles were filled from the neutral kit. Empty when the palette was complete. */
    neutralised: Array<keyof BrandPalette>;
    /** True when heading/body/docBody all came from the tenant record. */
    fontsFromTenant: boolean;
}
/** Pull the first `src="…"` out of a stored HTML snippet, or return the input if it is already a bare URL. */
export declare function extractImageUrl(v: unknown): string | undefined;
/** Infer a LogoAsset mime from a URL's extension. Undefined when unknown. */
export declare function mimeFromUrl(url: string): LogoAsset['mime'] | undefined;
/**
 * Normalise a tenant's raw branding record into a kit. Pure and synchronous.
 * `complete` is true only when all four palette roles, a light logo with
 * dimensions, and tenant-set fonts were present.
 */
export declare function normalise(tenant: TenantIdentity, branding: RawBranding): NormaliseResult;
export type MeasureFn = (url: string) => Promise<{
    widthPx: number;
    heightPx: number;
} | null>;
/**
 * Fill `logos.light` / `logos.dark` from legacy URLs using a caller-supplied
 * measurer (eq-field: an <img> + canvas; eq-shell functions: IHDR read; the
 * upload function: sharp/IHDR at upload time). Returns a new kit; the input
 * is not mutated. `complete` is recomputed.
 */
export declare function withMeasuredLogos(result: NormaliseResult, measure: MeasureFn): Promise<TenantBrandKit>;
/**
 * Adapter for the Shell→Service handoff JWT claims (brand_color + optional
 * brand_deep/ice/ink/doc_logo_url/doc_logo_dark_url). Produces the same
 * NormaliseResult shape so eq-service can keep the JWT as its transport.
 */
export declare function fromHandoffClaims(tenant: TenantIdentity, claims: Record<string, unknown>): NormaliseResult;
