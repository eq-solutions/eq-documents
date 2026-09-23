/**
 * IIFE entry for eq-field (vanilla JS, no bundler): window.EQDocuments.
 * Everything Field needs is bundled in (docx, jszip, contracts). xlsx is
 * deliberately excluded — Field has no spreadsheet writer, and exceljs
 * would triple the bundle.
 */
export * as brand from './brand/index.js';
export * as docx from './docx/index.js';
export * as template from './template/fillDocx.js';
export * as preflight from './preflight/check.js';
export { NEUTRAL_BRAND_KIT, BRAND_KIT_VERSION, validateTenantBrandKit, toHex6, isHex6 } from '@eq-solutions/contracts';
export declare const VERSION = "0.1.1";
