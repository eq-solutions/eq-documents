/**
 * docx.pack — create a kit-styled Document and serialise it in whichever
 * context we are in. `toBlob` is the browser path (eq-field, eq-shell client);
 * `toBuffer` is the server path (eq-service, Netlify functions).
 */
import { Document, Packer, type ISectionOptions } from 'docx'
import type { TenantBrandKit } from '@eq-solutions/contracts'
import { docxStyles } from './styles.js'
import { footer, pageA4 } from './primitives.js'

export interface CreateDocumentOptions {
  /** Document sections. `properties` defaults to A4/2 cm and `footers.default` to the kit footer unless given. */
  sections: ISectionOptions[]
  title?: string
  creator?: string
  description?: string
}

/** New Document with the kit's named styles, A4 page setup and legal footer applied to every section that doesn't override them. */
export function createDocument(kit: TenantBrandKit, opts: CreateDocumentOptions): Document {
  const kitFooter = footer(kit)
  return new Document({
    creator: opts.creator ?? kit.tenant.displayName,
    title: opts.title,
    description: opts.description,
    styles: docxStyles(kit),
    sections: opts.sections.map((s) => ({
      ...s,
      properties: s.properties ?? pageA4(),
      footers: s.footers ?? { default: kitFooter },
    })),
  })
}

export function toBlob(doc: Document): Promise<Blob> {
  return Packer.toBlob(doc)
}

export function toBuffer(doc: Document): Promise<Buffer> {
  return Packer.toBuffer(doc)
}

export function toUint8Array(doc: Document): Promise<Uint8Array> {
  return Packer.toBuffer(doc).then((b) => new Uint8Array(b))
}
