# @eq-solutions/documents

Shared branded-document engine for the EQ suite. One tenant brand kit in,
consistent Word / Excel documents out, for every app. Design and rationale:
`eq-context/eq/documents/branded-document-kit-design-2026-09-23.md`.

**Status: v0.1 — stage 1 of the design. No app consumes it yet.**

## What it does

| Module | Purpose |
|---|---|
| `brand.normalise(tenant, rawBranding)` | The one mapping from a tenant's raw `organisations.branding` jsonb to a `TenantBrandKit` (`@eq-solutions/contracts`). Understands today's legacy keys (`gateLogo` HTML, `hubLogo`, `gateLogoDark`) and tomorrow's structured `logos.*`. Missing values fall back to the **neutral grey** kit — never another tenant's brand. |
| `brand.withMeasuredLogos(result, measure)` | Fills logo dimensions from a caller-supplied measurer for legacy URLs. |
| `brand.fromHandoffClaims(tenant, claims)` | Adapter for eq-service's Shell→Service JWT brand claims. |
| `brand.contrastRatio / meetsAA / textOn` | WCAG helpers (moved from eq-shell `src/lib/contrast.ts`). |
| `docx.docxStyles(kit)` | Named Word styles from the kit: `DocTitle DocSubtitle DocH1 DocH2 DocBody DocSmall DocTableHead DocTableCell DocFooter`. The styles gallery in Word shows them; one edit restyles the file. |
| `docx.createDocument(kit, …)` | `docx` Document with those styles, A4/2 cm, and the kit's legal footer on every section. |
| `docx.masthead / footer / h1 / h2 / body / small / kvTable / dataTable / logoRun` | The shared vocabulary eq-field's `docx-builder.js` and eq-service's `lib/reports/*` each hand-roll today. Logo height is always derived from the **stored** aspect ratio. |
| `docx.toBlob / toBuffer / toUint8Array` | Browser and server serialisation. |
| `xlsx.xlsxTheme / applyHeaderRow / applyBodyRows / xlsxMasthead` | exceljs styling from the kit (lifted from eq-shell's compliance register). ESM only. |
| `template.fillDocx(bytes, bindings)` | Fill a tenant-authored letterhead `.docx` by content-control **tag** (body, headers, footers). No sentinel strings. Templates come from tenant storage, never an app repo. |
| `preflight.preflightDocx(kit, bytes)` | The six-line brand check on the bytes actually written: `Brand check: ✓ logo ✓ ratio ✓ palette ✓ fonts ✓ flat ✓ footer`. CI runs it on golden documents. |

## Consuming

| App | Context | Import |
|---|---|---|
| eq-shell, eq-service | bundled (Vite / Next) | `"@eq-solutions/documents": "github:eq-solutions/eq-documents#v0.1.0"` then `import { brand, docx } from '@eq-solutions/documents'` |
| eq-field | vanilla JS, no bundler | vendor `dist/documents.iife.js` to `scripts/vendor/eq-documents/` + a `VERSION` file, load it like `jszip.min.js`, use `window.EQDocuments`. Guard with a `documents-drift.yml` that diffs the vendored file against this repo at the pinned tag (same shape as eq-field's `tokens-drift.yml`). |
| eq-cards | Flutter | consumes the **contract only** (`@eq-solutions/contracts` `brand-kit.schema.json`). |

Both `dist/` outputs are committed; CI fails if a fresh build differs. The
IIFE bundles `docx`, `jszip` and the contract; it excludes `exceljs` on
purpose (Field has no spreadsheet writer; size budget 600 KB, currently ~480 KB).

## Rules baked in

- A kit is produced only by `brand.normalise` (or the canonical RPC that
  mirrors it). Consumers never assemble one by hand.
- Missing brand values render **neutral grey**, never EQ's or any other
  tenant's colours. A grey document is a visible "finish your brand card".
- No real tenant's values live in this repo. Tests use synthetic tenants.
- Logo height comes from stored dimensions. Nothing re-measures at render time.
- Flat: no gradients, shadows or effects. `preflight` fails on them.

## Developing

```sh
npm install
npm run typecheck
npm run build      # dist/documents.esm.js + dist/documents.iife.js + dist/types — commit them
npm test           # node:test against dist/, incl. golden documents + IIFE smoke test
```

Bump the version + tag `vX.Y.Z` on every change consumers should pick up;
the Release workflow attaches `dist/` to the GitHub release.

© 2026 CDC Solutions Pty Ltd. Proprietary and confidential.
