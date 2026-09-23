// Dual build.
//  dist/documents.esm.js  — for bundled consumers (eq-shell / eq-service). docx,
//                           exceljs, jszip and @eq-solutions/contracts stay
//                           external so the app's own copies are used.
//  dist/documents.iife.js — for eq-field (vanilla JS, no bundler). Everything
//                           it needs is bundled in and exposed as
//                           window.EQDocuments. xlsx is NOT in the IIFE —
//                           Field has no spreadsheet writer today and exceljs
//                           would triple the size.
// Both outputs are committed (eq-roles precedent: github: installs don't run
// build); CI fails if a rebuild differs from what's committed.
import { build } from 'esbuild'

const shared = { logLevel: 'info', sourcemap: true, target: ['es2020'], legalComments: 'none' }

await build({
  ...shared,
  entryPoints: ['src/index.ts'],
  outfile: 'dist/documents.esm.js',
  format: 'esm',
  platform: 'neutral',
  bundle: true,
  external: ['docx', 'exceljs', 'jszip', '@eq-solutions/contracts', '@eq-solutions/contracts/*'],
})

await build({
  ...shared,
  entryPoints: ['src/index.iife.ts'],
  outfile: 'dist/documents.iife.js',
  sourcemap: false,
  format: 'iife',
  globalName: 'EQDocuments',
  platform: 'browser',
  bundle: true,
  minify: true,
  define: { 'process.env.NODE_ENV': '"production"' },
})
console.log('built dist/documents.esm.js + dist/documents.iife.js')
