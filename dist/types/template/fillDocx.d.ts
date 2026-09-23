export interface FillResult {
    bytes: Uint8Array;
    /** Tags that were found and replaced. */
    filled: string[];
    /** Bindings whose tag did not exist in the template. */
    unmatched: string[];
}
export declare function escapeXml(s: string): string;
/**
 * Replace the content of every SDT with the given tag in a WordprocessingML
 * string. Exported for unit tests; consumers use fillDocx().
 */
export declare function replaceSdtByTag(xml: string, tag: string, text: string): {
    xml: string;
    count: number;
};
/** All SDT tags present in a document.xml string, in order of appearance. */
export declare function listSdtTags(xml: string): string[];
/**
 * Fill a .docx (bytes) by SDT tag. Touches word/document.xml plus any
 * header/footer parts, so a tag in a letterhead header is filled too.
 */
export declare function fillDocx(template: Uint8Array | ArrayBuffer | Blob, bindings: Record<string, string>): Promise<FillResult>;
