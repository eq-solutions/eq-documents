// Synthetic tenants only. No real tenant's values live in this repo — a real
// tenant's brand is that tenant's row in canonical, not test data.
import { NEUTRAL_BRAND_KIT } from '@eq-solutions/contracts'

export const acmeTenant = { id: '00000000-0000-0000-0000-00000000acc1', slug: 'acme', name: 'Acme', legalName: 'Acme Pty Ltd' }

/** Structured (post-migration) branding row. */
export const acmeBranding = {
  palette: { primary: '1F4E79', deep: '2E75B6', ice: 'EAF1FB', ink: '1B1B24', accent: 'C0504D' },
  logos: { light: { url: 'https://example.test/acme/logo.png', widthPx: 2000, heightPx: 723, mime: 'image/png' } },
  fonts: { heading: 'Roboto', body: 'Roboto', docBody: 'Calibri' },
  legal: { abn: '11 111 111 111', address: '1 Test St, Sydney NSW 2000', phone: '(02) 0000 0000' },
  // auth-gate flags that must be ignored
  rememberMeDays: 7,
  coreOnly: true,
}

/** Legacy (pre-migration) branding row, shaped like canonical today. */
export const legacyBranding = {
  palette: { primary: '1F4E79', deep: '2E75B6', ice: 'EAF1FB', ink: '1B1B24' },
  hubLogo: 'https://example.test/acme/hub.png',
  gateLogo: '<img src="https://example.test/acme/doc-logo.png" alt="Acme" style="height:64px;width:auto" />',
  gateLogoDark: 'https://example.test/acme/doc-logo-dark.png',
  orgName: 'Acme Field',
}

export const acmeKit = {
  kitVersion: 1,
  tenant: { id: acmeTenant.id, slug: 'acme', legalName: 'Acme Pty Ltd', displayName: 'Acme' },
  palette: acmeBranding.palette,
  logos: acmeBranding.logos,
  fonts: acmeBranding.fonts,
  legal: acmeBranding.legal,
  policy: { flat: true },
  complete: true,
}

export const neutralKit = { ...NEUTRAL_BRAND_KIT, tenant: { id: 'n', slug: 'nobody', legalName: 'Nobody Pty Ltd', displayName: 'Nobody' } }

// --- Guard fixtures: eq-context/eq/documents/auto-applied-layout-design-2026-09-27.md §2 ---

/** Wordmark-shaped logo (very wide, short) — masthead()'s width guard exists for exactly this shape. */
export const wordmarkKit = {
  ...acmeKit,
  tenant: { ...acmeKit.tenant, id: 'wordmark-1', slug: 'wordmark', legalName: 'Wordmark Co Pty Ltd', displayName: 'Wordmark Co' },
  logos: { light: { url: 'https://example.test/wordmark/logo.png', widthPx: 3000, heightPx: 400, mime: 'image/png' } },
}

/** Primary is near-white — too pale to read as a header fill. The header-fill intensity guard should fall back to palette.deep. */
export const paleTenantKit = {
  ...acmeKit,
  tenant: { ...acmeKit.tenant, id: 'pale-1', slug: 'pale', legalName: 'Pale Co Pty Ltd', displayName: 'Pale Co' },
  palette: { primary: 'F5F6F8', deep: '2E75B6', ice: 'EAF1FB', ink: '1B1B24', accent: 'C0504D' },
}

/** Primary is near-black — reads as an oversized black bar. Same guard, opposite band edge. */
export const inkTenantKit = {
  ...acmeKit,
  tenant: { ...acmeKit.tenant, id: 'ink-1', slug: 'inkco', legalName: 'Ink Co Pty Ltd', displayName: 'Ink Co' },
  palette: { primary: '050505', deep: '2E75B6', ice: 'EAF1FB', ink: '1B1B24', accent: 'C0504D' },
}

/** Heading font nobody ships by default — the font-fallback safety net's case. docBody stays a safe font (already gated upstream by DOC_BODY_SAFE_FONTS). */
export const unusualFontKit = {
  ...acmeKit,
  tenant: { ...acmeKit.tenant, id: 'display-font-1', slug: 'display-font', legalName: 'Display Font Co Pty Ltd', displayName: 'Display Font Co' },
  fonts: { heading: 'Brush Script MT', body: 'Brush Script MT', docBody: 'Calibri' },
}

/** 1×1 transparent PNG. */
export const PNG_1x1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64',
)
