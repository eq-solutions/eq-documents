/** WCAG 2.x contrast helpers on bare 6-digit hex. Moved here from eq-shell src/lib/contrast.ts so every consumer shares one implementation. */

export function hexToRgb(hex: string): [number, number, number] {
  const s = hex.replace(/^#/, '')
  return [parseInt(s.slice(0, 2), 16), parseInt(s.slice(2, 4), 16), parseInt(s.slice(4, 6), 16)]
}

function channel(c: number): number {
  const v = c / 255
  return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)
}

/** Relative luminance, 0 (black) … 1 (white). */
export function luminance(hex: string): number {
  const [r, g, b] = hexToRgb(hex)
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)
}

/** Contrast ratio 1 … 21 between two hex colours. Order-independent. */
export function contrastRatio(a: string, b: string): number {
  const la = luminance(a)
  const lb = luminance(b)
  const [hi, lo] = la > lb ? [la, lb] : [lb, la]
  return (hi + 0.05) / (lo + 0.05)
}

/** True when `fg` on `bg` meets WCAG AA for normal text (4.5:1). */
export function meetsAA(fg: string, bg: string): boolean {
  return contrastRatio(fg, bg) >= 4.5
}

/** True when `fg` on `bg` meets WCAG AAA for normal text (7:1). */
export function meetsAAA(fg: string, bg: string): boolean {
  return contrastRatio(fg, bg) >= 7
}

/** Pick white or the kit's ink for text on a given fill, whichever contrasts more. */
export function textOn(fillHex: string, inkHex: string): string {
  return contrastRatio('FFFFFF', fillHex) >= contrastRatio(inkHex, fillHex) ? 'FFFFFF' : inkHex
}
