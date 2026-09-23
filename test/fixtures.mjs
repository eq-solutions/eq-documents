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

/** 1×1 transparent PNG. */
export const PNG_1x1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64',
)
