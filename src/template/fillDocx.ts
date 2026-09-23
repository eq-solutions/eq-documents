/**
 * template.fillDocx — fill a tenant-authored .docx letterhead template by
 * Word content-control (SDT) tag. The one idea worth keeping from eq-shell's
 * quoteDocGenerator.ts, generalised: tags only, no sentinel-string replace,
 * no hardcoded rep blocks, template bytes supplied by the caller (from
 * tenant storage, never from an app repo's public/ folder).
 *
 * A binding replaces the whole <w:sdtContent> of every SDT whose <w:tag
 * w:val="…"> matches, with a single run that inherits the first run's
 * properties (so the template's own font/colour survive). Repeating tables
 * are out of scope — that is what the code-defined base documents are for.
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
