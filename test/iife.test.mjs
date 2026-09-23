import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import vm from 'node:vm'

// Load dist/documents.iife.js the way eq-field will: a plain <script>, no
// module system, exposing window.EQDocuments. Node has Blob/TextEncoder/etc.
test('IIFE bundle loads standalone and exposes the Field surface', async () => {
  const src = readFileSync(new URL('../dist/documents.iife.js', import.meta.url), 'utf8')
  const window = { Blob, TextEncoder, TextDecoder, setTimeout, clearTimeout, Uint8Array, ArrayBuffer, Promise, console }
  window.window = window
  window.self = window
  window.globalThis = window
  vm.runInNewContext(src, window)
  const EQ = window.EQDocuments
  assert.ok(EQ, 'window.EQDocuments defined')
  assert.equal(EQ.VERSION, '0.1.1')
  for (const k of ['brand', 'docx', 'template', 'preflight', 'NEUTRAL_BRAND_KIT', 'validateTenantBrandKit']) assert.ok(k in EQ, k)
  assert.equal('xlsx' in EQ, false, 'exceljs is not in the Field bundle')
  const r = EQ.brand.normalise({ id: 'x', slug: 'acme' }, { palette: { primary: '1F4E79', deep: '2E75B6', ice: 'EAF1FB', ink: '1B1B24' } })
  assert.equal(r.kit.palette.primary, '1F4E79')
  const doc = EQ.docx.createDocument(r.kit, { sections: [{ children: [EQ.docx.h1('Hello'), EQ.docx.body('World')] }] })
  const blob = await EQ.docx.toBlob(doc)
  assert.ok(blob.size > 1000)
})

test('IIFE bundle size stays under the Field budget (600 KB)', () => {
  const size = readFileSync(new URL('../dist/documents.iife.js', import.meta.url)).length
  assert.ok(size < 600 * 1024, `${size} bytes`)
})
