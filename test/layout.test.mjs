import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import { docx, preflight } from '../dist/documents.esm.js'
import { acmeKit, neutralKit } from './fixtures.mjs'

async function layoutDoc(kit) {
  const doc = docx.createDocument(kit, {
    title: 'Report',
    pageNumbers: true,
    header: { left: 'Maintenance Check Report', right: 'Site X' },
    sections: [
      {
        children: [
          docx.masthead(kit, { title: 'Report', subtitle: 'Site X' }),
          ...docx.toc(kit, [
            { label: 'Asset A', anchor: 'asset_a' },
            { label: 'Asset B', anchor: 'asset_b', indent: true },
          ]),
          docx.kpiRow(kit, [
            { label: 'Pass rate', value: '100%', sub: '4 / 4 tasks', status: 'pass' },
            { label: 'Assets', value: '04' },
            { label: 'Outstanding', value: '2', status: 'fail' },
          ]),
          docx.h1('Asset A', { bookmark: 'asset_a' }),
          docx.h1('Asset B', { bookmark: 'asset_b', pageBreakBefore: true }),
          docx.dataTable(kit, {
            head: ['Item', 'Result'],
            rows: [
              ['Greasing', { text: 'Pass', status: 'pass' }],
              ['Racking', { text: 'Fail', status: 'fail' }],
              ['Comms', { text: 'N/A', status: 'warn' }],
              ['Plain', 'x'],
            ],
          }),
        ],
      },
    ],
  })
  return docx.toUint8Array(doc)
}

test('layout: bookmarks, internal links, page break, header and page numbers are emitted', async () => {
  const zip = await JSZip.loadAsync(await layoutDoc(acmeKit))
  const document = await zip.file('word/document.xml').async('string')
  assert.ok(document.includes('w:bookmarkStart') && document.includes('w:name="asset_a"'), 'bookmark asset_a')
  assert.ok(document.includes('w:anchor="asset_b"'), 'internal link to asset_b')
  assert.equal((document.match(/w:pageBreakBefore/g) ?? []).length, 1, 'exactly one page break')
  const header = await zip.file('word/header1.xml').async('string')
  assert.ok(header.includes('Maintenance Check Report') && header.includes('Site X'))
  const footer = await zip.file('word/footer1.xml').async('string')
  assert.ok(footer.includes('PAGE') && footer.includes('NUMPAGES'))
})

test('layout: status cells are tinted; kpi figures use fixed result colours', async () => {
  const zip = await JSZip.loadAsync(await layoutDoc(acmeKit))
  const document = await zip.file('word/document.xml').async('string')
  for (const fill of Object.values(docx.STATUS_TINT)) assert.ok(document.includes(`w:fill="${fill}"`), `tint ${fill}`)
  assert.ok(document.includes('w:val="16A34A"'), 'kpi pass colour')
  assert.ok(document.includes('w:val="DC2626"'), 'kpi fail colour')
})

test('layout: document with status colours, kpi tiles and header passes preflight', async () => {
  for (const kit of [acmeKit, neutralKit]) {
    const bytes = await layoutDoc(kit)
    const facts = await preflight.extractFacts(bytes)
    const result = preflight.preflight(kit, { ...facts, logoSources: kit.logos.light ? [kit.logos.light.url] : [] })
    assert.equal(result.ok, true, result.line)
  }
})

test('layout: no header option means no header part', async () => {
  const doc = docx.createDocument(acmeKit, { sections: [{ children: [docx.body('x')] }] })
  const zip = await JSZip.loadAsync(await docx.toUint8Array(doc))
  assert.equal(zip.file('word/header1.xml'), null)
})

test('dataTable: fixed layout with an explicit column grid that fills the content width', async () => {
  const doc = docx.createDocument(acmeKit, {
    sections: [{ children: [docx.dataTable(acmeKit, { head: ['A', 'B', 'C'], widths: [50, 25, 25], rows: [['x', 'y', 'z']] })] }],
  })
  const zip = await JSZip.loadAsync(await docx.toUint8Array(doc))
  const xml = await zip.file('word/document.xml').async('string')
  assert.ok(xml.includes('w:tblLayout w:type="fixed"'), 'fixed layout')
  const cols = [...xml.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => Number(m[1]))
  assert.equal(cols.length, 3)
  assert.ok(Math.abs(cols.reduce((a, b) => a + b, 0) - 9638) <= 3, 'grid sums to A4 content width')
  assert.ok(cols[0] > cols[1] * 1.9, 'first column is ~2x the others')
})

test('dataTable: progress cell renders a proportional bar plus the label, preflight-clean', async () => {
  const doc = docx.createDocument(acmeKit, {
    sections: [{ children: [docx.dataTable(acmeKit, { head: ['Task', 'Done'], rows: [['a', { text: '3/4', progress: { done: 3, total: 4 } }], ['b', { text: '0/4', progress: { done: 0, total: 4 } }]] })] }],
  })
  const bytes = await docx.toUint8Array(doc)
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml').async('string')
  const nbsp = (xml.match(/ +/g) ?? []).map((m) => m.length)
  assert.deepEqual(nbsp, [15, 5, 20], '3/4 -> 15 filled + 5 track; 0/4 -> 20 track')
  assert.ok(xml.includes('3/4') && xml.includes('0/4'))
  const facts = await preflight.extractFacts(bytes)
  const result = preflight.preflight(acmeKit, { ...facts, logoSources: [acmeKit.logos.light.url] })
  assert.equal(result.ok, true, result.line)
})
