import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import { template } from '../dist/documents.esm.js'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function sdt(tag, inner) {
  return `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
}

/** Minimal letterhead-style template: a block SDT in the body, an inline SDT in a paragraph, an SDT in the header. */
async function tinyTemplate() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>` +
    sdt('ClientName', `<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="112233"/></w:rPr><w:t>Click to enter client</w:t></w:r></w:p>`) +
    `<w:p><w:r><w:t>Dear </w:t></w:r>` + sdt('DearName', `<w:r><w:rPr><w:i/></w:rPr><w:t>name</w:t></w:r>`) + `<w:r><w:t>,</w:t></w:r></w:p>` +
    sdt('Untouched', `<w:p><w:r><w:t>leave me</w:t></w:r></w:p>`) +
    `</w:body></w:document>`)
  zip.file('word/header1.xml', `<?xml version="1.0"?><w:hdr ${W}>` + sdt('QuoteNumber', `<w:p><w:r><w:t>Q-0000</w:t></w:r></w:p>`) + `</w:hdr>`)
  return zip.generateAsync({ type: 'uint8array' })
}

test('fills block, inline and header SDTs by tag; keeps run properties; escapes XML; reports unmatched', async () => {
  const bytes = await tinyTemplate()
  const r = await template.fillDocx(bytes, { ClientName: 'Data Centre Client A', DearName: 'Sam', QuoteNumber: 'Q-1234', NoSuchTag: 'x', Multi: 'a\nb' })
  assert.deepEqual([...r.filled].sort(), ['ClientName', 'DearName', 'QuoteNumber'])
  assert.deepEqual(r.unmatched.sort(), ['Multi', 'NoSuchTag'])
  const zip = await JSZip.loadAsync(r.bytes)
  const doc = await zip.file('word/document.xml').async('string')
  const hdr = await zip.file('word/header1.xml').async('string')
  assert.ok(doc.includes('<w:t xml:space="preserve">Data Centre Client A</w:t>'))
  assert.ok(doc.includes('<w:rPr><w:b/><w:color w:val="112233"/></w:rPr><w:t xml:space="preserve">Data Centre Client A</w:t>'), 'first run rPr preserved')
  assert.ok(doc.includes('<w:pPr><w:jc w:val="left"/></w:pPr>'), 'block SDT keeps paragraph props')
  assert.ok(doc.includes('<w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Sam</w:t>'), 'inline SDT stays inline')
  assert.ok(!doc.includes('Click to enter client'))
  assert.ok(doc.includes('leave me'))
  assert.ok(hdr.includes('Q-1234'))
})

test('escapes XML special characters and splits newlines into breaks', () => {
  const xml = sdt('T', `<w:p><w:r><w:t>x</w:t></w:r></w:p>`)
  const r = template.replaceSdtByTag(xml, 'T', 'A & B <C>\nline 2')
  assert.equal(r.count, 1)
  assert.ok(r.xml.includes('A &amp; B &lt;C&gt;</w:t><w:br/><w:t xml:space="preserve">line 2'))
})

test('listSdtTags', () => {
  assert.deepEqual(template.listSdtTags(sdt('A', '') + sdt('B', '')), ['A', 'B'])
})
