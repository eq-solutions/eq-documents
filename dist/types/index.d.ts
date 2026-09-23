/**
 * @eq-solutions/documents — shared branded-document engine for the EQ suite.
 * ESM entry (bundled consumers). See src/index.iife.ts for eq-field's build.
 */
export * as brand from './brand/index.js';
export * as docx from './docx/index.js';
export * as xlsx from './xlsx/theme.js';
export * as template from './template/fillDocx.js';
export * as preflight from './preflight/check.js';
export { NEUTRAL_BRAND_KIT, BRAND_KIT_VERSION, validateTenantBrandKit, toHex6, isHex6 } from '@eq-solutions/contracts';
export type { TenantBrandKit, BrandPalette, BrandFonts, BrandLegal, LogoAsset } from '@eq-solutions/contracts';
