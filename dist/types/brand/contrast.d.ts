/** WCAG 2.x contrast helpers on bare 6-digit hex. Moved here from eq-shell src/lib/contrast.ts so every consumer shares one implementation. */
export declare function hexToRgb(hex: string): [number, number, number];
/** Relative luminance, 0 (black) … 1 (white). */
export declare function luminance(hex: string): number;
/** Contrast ratio 1 … 21 between two hex colours. Order-independent. */
export declare function contrastRatio(a: string, b: string): number;
/** True when `fg` on `bg` meets WCAG AA for normal text (4.5:1). */
export declare function meetsAA(fg: string, bg: string): boolean;
/** True when `fg` on `bg` meets WCAG AAA for normal text (7:1). */
export declare function meetsAAA(fg: string, bg: string): boolean;
/** Pick white or the kit's ink for text on a given fill, whichever contrasts more. */
export declare function textOn(fillHex: string, inkHex: string): string;
