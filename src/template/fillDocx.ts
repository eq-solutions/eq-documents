/**
 * template.fillDocx — fill a tenant-authored .docx letterhead template by
 * Word content-control (SDT) tag. The one idea worth keeping from eq-shell's
 * quoteDocGenerator.ts, generalised: tags only, no sentinel-string replace,
 * no hardcoded rep blocks, template bytes supplied by the caller (from
 * tenant storage, never from an app repo's public/ folder).
 *
 * A binding replaces the whole <w:sdtContent> of every SDT whose <w:tag
 * w:val="…"> matches, with a single run that inherits the first run's
 * properties (so the template's own font/colour survive).
 *
 * `fillDocx` alone fills a fixed set of tags — it cannot grow a table. For a
 * variable-row section (e.g. quote line items), use `fillDocxWithLoops` /
 * `replaceLoopByTag` below: one templated row, repeated per data row. A
 * document whose table *shape* (not just row count) varies still belongs in
 * code (docx.* + primitives), not a letterhead template — loops here are
 * deliberately one level, no nesting, no conditionals.
 */
import JSZip from 'jszip'

export interface FillResult {
  bytes: Uint8Array
  /** Tags that were found and replaced. */
  filled: string[]
  /** Bindings whose tag did not exist in the template. */
  unmatched: string[]
}

export function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

/** Split on newlines into <w:t> runs separated by <w:br/>, inside one run. */
function runXml(text: string, rPr: string): string {
  const lines = text.split(/\r?\n/)
  const inner = lines.map((l) => `<w:t xml:space="preserve">${escapeXml(l)}</w:t>`).join('<w:br/>')
  return `<w:r>${rPr}${inner}</w:r>`
}

/**
 * Replace the content of every SDT with the given tag in a WordprocessingML
 * string. Exported for unit tests; consumers use fillDocx().
 */
export function replaceSdtByTag(xml: string, tag: string, text: string): { xml: string; count: number } {
  let count = 0
  // Match each <w:sdt> … </w:sdt> block lazily; SDTs may nest in tables but
  // not inside each other in any template we own, so a non-greedy scan is safe.
  const re = /<w:sdt>([\s\S]*?)<\/w:sdt>/g
  const out = xml.replace(re, (whole, inner: string) => {
    const tagMatch = inner.match(/<w:tag\s+w:val="([^"]*)"/)
    if (!tagMatch || tagMatch[1] !== tag) return whole
    const contentMatch = inner.match(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/)
    if (!contentMatch) return whole
    const content = contentMatch[1]
    // Preserve paragraph wrapper if the SDT is block-level, and the first run's properties.
    const rPr = (content.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) ?? [''])[0]
    const pPr = (content.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) ?? [''])[0]
    const isBlock = /<w:p[\s>]/.test(content)
    const replacement = isBlock ? `<w:p>${pPr}${runXml(text, rPr)}</w:p>` : runXml(text, rPr)
    count++
    return whole.replace(contentMatch[0], `<w:sdtContent>${replacement}</w:sdtContent>`)
  })
  return { xml: out, count }
}

/** All SDT tags present in a document.xml string, in order of appearance. */
export function listSdtTags(xml: string): string[] {
  const tags: string[] = []
  const re = /<w:tag\s+w:val="([^"]*)"/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml))) tags.push(m[1])
  return tags
}

/**
 * Fill a .docx (bytes) by SDT tag. Touches word/document.xml plus any
 * header/footer parts, so a tag in a letterhead header is filled too.
 */
export async function fillDocx(template: Uint8Array | ArrayBuffer | Blob, bindings: Record<string, string>): Promise<FillResult> {
  const zip = await JSZip.loadAsync(template as Uint8Array)
  const parts = Object.keys(zip.files).filter((p) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(p))
  const filled = new Set<string>()
  for (const part of parts) {
    let xml = await zip.file(part)!.async('string')
    for (const [tag, value] of Object.entries(bindings)) {
      const r = replaceSdtByTag(xml, tag, value ?? '')
      if (r.count > 0) filled.add(tag)
      xml = r.xml
    }
    zip.file(part, xml)
  }
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  const unmatched = Object.keys(bindings).filter((t) => !filled.has(t))
  return { bytes, filled: [...filled], unmatched }
}

export interface LoopFillResult extends FillResult {
  /** Loop tags that were found and expanded. */
  loopsFilled: string[]
  /** Loop tags in `loops` whose tag did not exist in the template. */
  loopsUnmatched: string[]
}

/**
 * Find the `<w:sdt>…</w:sdt>` span whose OWN `<w:sdtPr><w:tag w:val="tag">`
 * matches — not a nested one. A loop's content contains per-item SDTs
 * (nested), so unlike `replaceSdtByTag`'s single non-greedy regex (correct
 * only for non-nested SDTs, per that function's own comment), this needs a
 * depth-aware scan to find the loop wrapper's *matching* close tag rather
 * than the first `</w:sdt>` encountered (which would belong to the first
 * nested per-item SDT instead).
 */
function findNamedSdtSpan(
  xml: string,
  tag: string
): { start: number; end: number; contentStart: number; contentEnd: number } | null {
  const tagNeedle = `<w:tag w:val="${tag}"`
  let searchFrom = 0
  while (true) {
    const tagIdx = xml.indexOf(tagNeedle, searchFrom)
    if (tagIdx === -1) return null
    const start = xml.lastIndexOf('<w:sdt>', tagIdx)
    // <w:sdtPr> (holding <w:tag>) is always the first child of <w:sdt>, before
    // any nested <w:sdt> can appear — so the nearest preceding "<w:sdt>" is
    // always this tag's own wrapper, never an ancestor's.
    if (start === -1) {
      searchFrom = tagIdx + tagNeedle.length
      continue
    }
    const boundaryRe = /<w:sdt>|<\/w:sdt>/g
    boundaryRe.lastIndex = start
    let depth = 0
    let end = -1
    let m: RegExpExecArray | null
    while ((m = boundaryRe.exec(xml))) {
      if (m[0] === '<w:sdt>') depth++
      else {
        depth--
        if (depth === 0) {
          end = m.index + m[0].length
          break
        }
      }
    }
    if (end === -1) return null
    const contentOpen = xml.indexOf('<w:sdtContent>', start)
    const contentCloseTag = '</w:sdtContent>'
    const sdtCloseStart = end - '</w:sdt>'.length
    const contentClose = xml.lastIndexOf(contentCloseTag, sdtCloseStart)
    if (contentOpen === -1 || contentClose === -1 || contentOpen >= contentClose) return null
    return { start, end, contentStart: contentOpen + '<w:sdtContent>'.length, contentEnd: contentClose }
  }
}

/**
 * Expand a repeating-row loop SDT: authored in Word as one table row wrapped
 * in a content control (Insert > Controls > Rich Text, tag it e.g.
 * "Loop:LineItems"), with per-cell content controls inside it (tag e.g.
 * "Item.Description", "Item.Total"). One row in the template becomes N rows,
 * one per entry in `items`, each filled by `replaceSdtByTag` against that
 * item's own bindings. A non-table loop (SDT content with no <w:tr>) repeats
 * the whole block instead — same mechanism, without the row unwrap.
 */
export function replaceLoopByTag(xml: string, tag: string, items: Array<Record<string, string>>): { xml: string; count: number } {
  const span = findNamedSdtSpan(xml, tag)
  if (!span) return { xml, count: 0 }
  const content = xml.slice(span.contentStart, span.contentEnd)
  const rowMatch = content.match(/<w:tr[\s>][\s\S]*<\/w:tr>/)
  const templateBlock = rowMatch ? rowMatch[0] : content
  const rows = items
    .map((item) => {
      let rowXml = templateBlock
      for (const [itemTag, value] of Object.entries(item)) {
        rowXml = replaceSdtByTag(rowXml, itemTag, value ?? '').xml
      }
      return rowXml
    })
    .join('')
  // Drop the loop's own SDT wrapper entirely — the repeated rows sit directly
  // in the table (or body) in its place, so Word sees N ordinary rows, not
  // one content control containing N rows (which OOXML doesn't allow for a
  // row-level control anyway).
  const out = xml.slice(0, span.start) + rows + xml.slice(span.end)
  return { xml: out, count: 1 }
}

/**
 * fillDocx plus one or more repeating-row loops. Loops expand first (so a
 * loop's own per-item SDTs never collide with top-level `bindings` of the
 * same tag name), then ordinary `bindings` fill the rest of the document as
 * usual.
 */
export async function fillDocxWithLoops(
  template: Uint8Array | ArrayBuffer | Blob,
  bindings: Record<string, string>,
  loops: Record<string, Array<Record<string, string>>>
): Promise<LoopFillResult> {
  const zip = await JSZip.loadAsync(template as Uint8Array)
  const parts = Object.keys(zip.files).filter((p) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(p))
  const loopsFilled = new Set<string>()
  for (const part of parts) {
    let xml = await zip.file(part)!.async('string')
    for (const [loopTag, items] of Object.entries(loops)) {
      const r = replaceLoopByTag(xml, loopTag, items)
      if (r.count > 0) loopsFilled.add(loopTag)
      xml = r.xml
    }
    zip.file(part, xml)
  }
  const rebuilt = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  const base = await fillDocx(rebuilt, bindings)
  const loopsUnmatched = Object.keys(loops).filter((t) => !loopsFilled.has(t))
  return { ...base, loopsFilled: [...loopsFilled], loopsUnmatched }
}
