import { test } from 'node:test'
import assert from 'node:assert/strict'
import { brand, validateTenantBrandKit, NEUTRAL_BRAND_KIT } from '../dist/documents.esm.js'
import { acmeTenant, acmeBranding, legacyBranding } from './fixtures.mjs'

test('structured branding → complete, valid kit; auth-gate flags ignored', () => {
  const r = brand.normalise(acmeTenant, acmeBranding)
  assert.equal(validateTenantBrandKit(r.kit).ok, true)
  assert.equal(r.kit.complete, true)
  assert.deepEqual(r.neutralised, [])
  assert.equal(r.kit.palette.accent, 'C0504D')
  assert.equal(r.kit.logos.light.heightPx, 723)
  assert.equal(r.kit.tenant.legalName, 'Acme Pty Ltd')
  assert.equal('rememberMeDays' in r.kit, false)
})

test('legacy branding → valid but incomplete kit, logo URLs surfaced for measuring', () => {
  const r = brand.normalise(acmeTenant, legacyBranding)
  assert.equal(validateTenantBrandKit(r.kit).ok, true)
  assert.equal(r.kit.complete, false)
  assert.deepEqual(r.kit.logos, {})
  assert.equal(r.legacyLogos.light, 'https://example.test/acme/doc-logo.png', 'gateLogo <img> src wins over hubLogo')
  assert.equal(r.legacyLogos.dark, 'https://example.test/acme/doc-logo-dark.png')
  assert.equal(r.kit.tenant.displayName, 'Acme Field', 'orgName is the display name')
  assert.equal(r.kit.fonts.docBody, NEUTRAL_BRAND_KIT.fonts.docBody)
})

test('withMeasuredLogos fills dimensions from the caller measurer and recomputes complete', async () => {
  const r = brand.normalise(acmeTenant, { ...legacyBranding, fonts: acmeBranding.fonts })
  const kit = await brand.withMeasuredLogos(r, async (url) => (url.endsWith('doc-logo.png') ? { widthPx: 1200, heightPx: 400 } : null))
  assert.equal(kit.logos.light.widthPx, 1200)
  assert.equal(kit.logos.light.mime, 'image/png')
  assert.equal(kit.logos.dark, undefined, 'unmeasurable logo is dropped, not faked')
  assert.equal(kit.complete, true)
  assert.equal(r.kit.logos.light, undefined, 'input not mutated')
})

test('missing or invalid palette roles fall back to NEUTRAL grey per role, never a brand', () => {
  const r = brand.normalise(acmeTenant, { palette: { primary: '#1F4E79', deep: 'not-a-colour', ice: 'EAF1FB' } })
  assert.equal(r.kit.palette.primary, '1F4E79')
  assert.equal(r.kit.palette.deep, NEUTRAL_BRAND_KIT.palette.deep)
  assert.equal(r.kit.palette.ink, NEUTRAL_BRAND_KIT.palette.ink)
  assert.deepEqual(r.neutralised, ['deep', 'ink'])
  assert.equal(r.kit.complete, false)
  // No EQ or any other brand hue sneaks in as a default.
  for (const hex of Object.values(r.kit.palette)) assert.notEqual(hex, '3DA8D8')
})

test('empty branding → entirely neutral kit', () => {
  const r = brand.normalise({ id: 'x', slug: 'blank' }, {})
  assert.equal(validateTenantBrandKit(r.kit).ok, true)
  assert.deepEqual(r.kit.palette, { ...NEUTRAL_BRAND_KIT.palette })
  assert.equal(r.kit.tenant.displayName, 'blank')
  assert.equal(r.kit.complete, false)
})

test('docBody outside the safe list is neutralised, heading/body kept', () => {
  const r = brand.normalise(acmeTenant, { ...acmeBranding, fonts: { heading: 'Roboto', body: 'Roboto', docBody: 'Roboto' } })
  assert.equal(r.kit.fonts.heading, 'Roboto')
  assert.equal(r.kit.fonts.docBody, 'Arial')
  assert.equal(r.kit.complete, false)
})

test('fromHandoffClaims maps the JWT brand claims', () => {
  const r = brand.fromHandoffClaims(acmeTenant, { brand_color: '#1F4E79', brand_deep: '2E75B6', brand_ice: 'EAF1FB', brand_ink: '1B1B24', brand_doc_logo_url: 'https://example.test/x.png' })
  assert.equal(r.kit.palette.primary, '1F4E79')
  assert.equal(r.legacyLogos.light, 'https://example.test/x.png')
})

test('extractImageUrl and mimeFromUrl', () => {
  assert.equal(brand.extractImageUrl('<div><img alt="x" src=\'https://a/b.webp\'></div>'), 'https://a/b.webp')
  assert.equal(brand.extractImageUrl('https://a/b.png'), 'https://a/b.png')
  assert.equal(brand.extractImageUrl(''), undefined)
  assert.equal(brand.mimeFromUrl('https://a/b.SVG?v=1'), 'image/svg+xml')
  assert.equal(brand.mimeFromUrl('https://a/b'), undefined)
})

test('contrast helpers', () => {
  assert.ok(Math.abs(brand.contrastRatio('000000', 'FFFFFF') - 21) < 0.01)
  assert.equal(brand.meetsAA('1B1B24', 'FFFFFF'), true)
  assert.equal(brand.textOn('EAF1FB', '1B1B24'), '1B1B24')
  assert.equal(brand.textOn('1F4E79', '1B1B24'), 'FFFFFF')
})
