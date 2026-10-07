/**
 * docx.pack — create a kit-styled Document and serialise it in whichever
 * context we are in. `toBlob` is the browser path (eq-field, eq-shell client);
 * `toBuffer` is the server path (eq-service, Netlify functions).
 */
import { Document, type ISectionOptions } from 'docx';
import type { TenantBrandKit } from '@eq-solutions/contracts';
import { type PageHeaderOptions } from './layout.js';
export interface CreateDocumentOptions {
    /** Document sections. `properties` defaults to A4/2 cm and `footers.default` to the kit footer unless given. */
    sections: ISectionOptions[];
    title?: string;
    creator?: string;
    description?: string;
    /** Append page-number fields (`Page N of M`) to the default kit footer. Default false — see `FooterOptions`. Ignored on any section that provides its own `footers`. */
    pageNumbers?: boolean;
    /** Running page header on every section that doesn't provide its own `headers`. Omit for none. */
    header?: PageHeaderOptions;
}
/** New Document with the kit's named styles, A4 page setup and legal footer applied to every section that doesn't override them. */
export declare function createDocument(kit: TenantBrandKit, opts: CreateDocumentOptions): Document;
export declare function toBlob(doc: Document): Promise<Blob>;
export declare function toBuffer(doc: Document): Promise<Buffer>;
export declare function toUint8Array(doc: Document): Promise<Uint8Array>;
