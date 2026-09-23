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
export interface LoopFillResult extends FillResult {
    /** Loop tags that were found and expanded. */
    loopsFilled: string[];
    /** Loop tags in `loops` whose tag did not exist in the template. */
    loopsUnmatched: string[];
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
export declare function replaceLoopByTag(xml: string, tag: string, items: Array<Record<string, string>>): {
    xml: string;
    count: number;
};
/**
 * fillDocx plus one or more repeating-row loops. Loops expand first (so a
 * loop's own per-item SDTs never collide with top-level `bindings` of the
 * same tag name), then ordinary `bindings` fill the rest of the document as
 * usual.
 */
export declare function fillDocxWithLoops(template: Uint8Array | ArrayBuffer | Blob, bindings: Record<string, string>, loops: Record<string, Array<Record<string, string>>>): Promise<LoopFillResult>;
