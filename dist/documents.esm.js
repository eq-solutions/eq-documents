var __defProp = Object.defineProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};

// src/brand/index.ts
var brand_exports = {};
__export(brand_exports, {
  contrastRatio: () => contrastRatio,
  extractImageUrl: () => extractImageUrl,
  fromHandoffClaims: () => fromHandoffClaims,
  hexToRgb: () => hexToRgb,
  luminance: () => luminance,
  meetsAA: () => meetsAA,
  meetsAAA: () => meetsAAA,
  mimeFromUrl: () => mimeFromUrl,
  normalise: () => normalise,
  textOn: () => textOn,
  withMeasuredLogos: () => withMeasuredLogos
});

// src/brand/normalise.ts
import {
  BRAND_KIT_VERSION,
  DOC_BODY_SAFE_FONTS,
  NEUTRAL_BRAND_KIT,
  toHex6
} from "@eq-solutions/contracts";
var LOGO_MIMES = ["image/png", "image/svg+xml", "image/webp", "image/jpeg"];
function asRecord(v) {
  return typeof v === "object" && v !== null && !Array.isArray(v) ? v : {};
}
function asString(v) {
  return typeof v === "string" && v.trim() !== "" ? v.trim() : void 0;
}
function extractImageUrl(v) {
  const s = asString(v);
  if (!s) return void 0;
  if (!s.includes("<")) return s;
  const m = s.match(/<img[^>]*\ssrc\s*=\s*["']([^"']+)["']/i);
  return m?.[1];
}
function mimeFromUrl(url) {
  const clean = url.split("?")[0].split("#")[0].toLowerCase();
  if (clean.endsWith(".png")) return "image/png";
  if (clean.endsWith(".svg")) return "image/svg+xml";
  if (clean.endsWith(".webp")) return "image/webp";
  if (clean.endsWith(".jpg") || clean.endsWith(".jpeg")) return "image/jpeg";
  if (clean.startsWith("data:image/png")) return "image/png";
  if (clean.startsWith("data:image/svg+xml")) return "image/svg+xml";
  return void 0;
}
function toLogoAsset(v) {
  const o = asRecord(v);
  const url = asString(o.url);
  const widthPx = typeof o.widthPx === "number" ? o.widthPx : Number(o.widthPx);
  const heightPx = typeof o.heightPx === "number" ? o.heightPx : Number(o.heightPx);
  const mime = asString(o.mime) ?? (url ? mimeFromUrl(url) : void 0);
  if (!url || !(widthPx > 0) || !(heightPx > 0) || !mime || !LOGO_MIMES.includes(mime)) return void 0;
  return { url, widthPx, heightPx, mime };
}
function normalisePalette(raw) {
  const o = asRecord(raw);
  const neutralised = [];
  const pick = (k) => {
    const hex = toHex6(o[k]);
    if (hex) return hex;
    neutralised.push(k);
    return NEUTRAL_BRAND_KIT.palette[k];
  };
  const palette = { primary: pick("primary"), deep: pick("deep"), ice: pick("ice"), ink: pick("ink") };
  const accent = toHex6(o.accent);
  if (accent) palette.accent = accent;
  return { palette, neutralised };
}
function normaliseFonts(raw) {
  const o = asRecord(raw);
  const heading = asString(o.heading);
  const body2 = asString(o.body);
  const docBodyRaw = asString(o.docBody);
  const docBody = docBodyRaw && DOC_BODY_SAFE_FONTS.includes(docBodyRaw) ? docBodyRaw : void 0;
  const fromTenant = Boolean(heading && body2 && docBody);
  return {
    fonts: {
      heading: heading ?? NEUTRAL_BRAND_KIT.fonts.heading,
      body: body2 ?? heading ?? NEUTRAL_BRAND_KIT.fonts.body,
      docBody: docBody ?? NEUTRAL_BRAND_KIT.fonts.docBody
    },
    fromTenant
  };
}
function normaliseLegal(raw) {
  const o = asRecord(raw);
  const legal = {};
  for (const k of ["abn", "address", "phone", "email", "web"]) {
    const v = asString(o[k]);
    if (v) legal[k] = v;
  }
  return legal;
}
function normalise(tenant, branding) {
  const b = asRecord(branding);
  const { palette, neutralised } = normalisePalette(b.palette);
  const { fonts, fromTenant: fontsFromTenant } = normaliseFonts(b.fonts);
  const legal = normaliseLegal(b.legal);
  const structured = asRecord(b.logos);
  const logos = {};
  const light = toLogoAsset(structured.light);
  const dark = toLogoAsset(structured.dark);
  const mark = toLogoAsset(structured.mark);
  if (light) logos.light = light;
  if (dark) logos.dark = dark;
  if (mark) logos.mark = mark;
  const legacyLogos = {};
  if (!light) {
    const url = extractImageUrl(b.gateLogo) ?? extractImageUrl(b.sidebarLogoHtml) ?? extractImageUrl(b.hubLogo);
    if (url) legacyLogos.light = url;
  }
  if (!dark) {
    const url = extractImageUrl(b.gateLogoDark);
    if (url) legacyLogos.dark = url;
  }
  const displayName = asString(tenant.name) ?? tenant.slug;
  const legalName = asString(asRecord(b.legal).legalName) ?? asString(tenant.legalName) ?? displayName;
  const kit = {
    kitVersion: BRAND_KIT_VERSION,
    tenant: { id: tenant.id, slug: tenant.slug, legalName, displayName },
    palette,
    logos,
    fonts,
    legal,
    policy: { flat: true },
    complete: neutralised.length === 0 && Boolean(light) && fontsFromTenant
  };
  return { kit, legacyLogos, neutralised, fontsFromTenant };
}
async function withMeasuredLogos(result, measure) {
  const logos = { ...result.kit.logos };
  for (const role of ["light", "dark"]) {
    const url = result.legacyLogos[role];
    if (!url || logos[role]) continue;
    const mime = mimeFromUrl(url);
    if (!mime) continue;
    const dims = await measure(url).catch(() => null);
    if (dims && dims.widthPx > 0 && dims.heightPx > 0) logos[role] = { url, widthPx: dims.widthPx, heightPx: dims.heightPx, mime };
  }
  const complete = result.neutralised.length === 0 && Boolean(logos.light) && result.fontsFromTenant;
  return { ...result.kit, logos, complete };
}
function fromHandoffClaims(tenant, claims) {
  const branding = {
    palette: { primary: claims.brand_color, deep: claims.brand_deep, ice: claims.brand_ice, ink: claims.brand_ink },
    gateLogo: claims.brand_doc_logo_url ?? claims.brand_logo_url,
    gateLogoDark: claims.brand_doc_logo_dark_url,
    fonts: claims.brand_fonts,
    legal: claims.brand_legal
  };
  return normalise(tenant, branding);
}

// src/brand/contrast.ts
function hexToRgb(hex) {
  const s = hex.replace(/^#/, "");
  return [parseInt(s.slice(0, 2), 16), parseInt(s.slice(2, 4), 16), parseInt(s.slice(4, 6), 16)];
}
function channel(c) {
  const v = c / 255;
  return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
}
function luminance(hex) {
  const [r, g, b] = hexToRgb(hex);
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}
function contrastRatio(a, b) {
  const la = luminance(a);
  const lb = luminance(b);
  const [hi, lo] = la > lb ? [la, lb] : [lb, la];
  return (hi + 0.05) / (lo + 0.05);
}
function meetsAA(fg, bg) {
  return contrastRatio(fg, bg) >= 4.5;
}
function meetsAAA(fg, bg) {
  return contrastRatio(fg, bg) >= 7;
}
function textOn(fillHex, inkHex) {
  return contrastRatio("FFFFFF", fillHex) >= contrastRatio(inkHex, fillHex) ? "FFFFFF" : inkHex;
}

// src/docx/index.ts
var docx_exports = {};
__export(docx_exports, {
  DOC_STYLE_IDS: () => DOC_STYLE_IDS,
  HAIRLINE: () => HAIRLINE,
  accentOf: () => accentOf,
  body: () => body,
  createDocument: () => createDocument,
  dataTable: () => dataTable,
  docxStyles: () => docxStyles,
  footer: () => footer,
  footerText: () => footerText,
  h1: () => h1,
  h2: () => h2,
  kvTable: () => kvTable,
  logoRun: () => logoRun,
  masthead: () => masthead,
  mutedInk: () => mutedInk,
  pageA4: () => pageA4,
  small: () => small,
  spacer: () => spacer,
  toBlob: () => toBlob,
  toBuffer: () => toBuffer,
  toUint8Array: () => toUint8Array
});

// src/docx/styles.ts
import { AlignmentType, BorderStyle } from "docx";
var DOC_STYLE_IDS = [
  "DocTitle",
  "DocSubtitle",
  "DocH1",
  "DocH2",
  "DocBody",
  "DocSmall",
  "DocTableHead",
  "DocTableCell",
  "DocFooter"
];
function accentOf(kit) {
  return kit.palette.accent ?? kit.palette.deep;
}
function mutedInk(kit) {
  const ink = kit.palette.ink;
  const mix = (i) => Math.round(parseInt(ink.slice(i, i + 2), 16) * 0.6 + 255 * 0.4);
  return [0, 2, 4].map((i) => mix(i).toString(16).padStart(2, "0")).join("").toUpperCase();
}
function docxStyles(kit) {
  const { primary, deep, ink } = kit.palette;
  const accent = accentOf(kit);
  const heading = kit.fonts.heading;
  const body2 = kit.fonts.docBody;
  return {
    default: {
      document: { run: { font: body2, size: 22, color: ink } }
    },
    paragraphStyles: [
      {
        id: "DocTitle",
        name: "Doc Title",
        basedOn: "Normal",
        next: "DocBody",
        quickFormat: true,
        run: { font: heading, size: 36, bold: true, color: primary },
        paragraph: { spacing: { before: 0, after: 120 } }
      },
      {
        id: "DocSubtitle",
        name: "Doc Subtitle",
        basedOn: "Normal",
        next: "DocBody",
        quickFormat: true,
        run: { font: heading, size: 22, color: deep },
        paragraph: {
          spacing: { before: 0, after: 240 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: accent, space: 4 } }
        }
      },
      {
        id: "DocH1",
        name: "Doc Heading 1",
        basedOn: "Normal",
        next: "DocBody",
        quickFormat: true,
        run: { font: heading, size: 28, bold: true, color: primary },
        paragraph: { spacing: { before: 320, after: 120 }, keepNext: true, outlineLevel: 0 }
      },
      {
        id: "DocH2",
        name: "Doc Heading 2",
        basedOn: "Normal",
        next: "DocBody",
        quickFormat: true,
        run: { font: heading, size: 24, bold: true, color: deep },
        paragraph: { spacing: { before: 240, after: 80 }, keepNext: true, outlineLevel: 1 }
      },
      {
        id: "DocBody",
        name: "Doc Body",
        basedOn: "Normal",
        quickFormat: true,
        run: { font: body2, size: 22, color: ink },
        paragraph: { spacing: { after: 120, line: 276 } }
      },
      {
        id: "DocSmall",
        name: "Doc Small",
        basedOn: "DocBody",
        run: { size: 16, color: mutedInk(kit) },
        paragraph: { spacing: { after: 60 } }
      },
      {
        id: "DocTableHead",
        name: "Doc Table Head",
        basedOn: "DocBody",
        run: { size: 18, bold: true, color: "FFFFFF" },
        paragraph: { spacing: { before: 40, after: 40 } }
      },
      {
        id: "DocTableCell",
        name: "Doc Table Cell",
        basedOn: "DocBody",
        run: { size: 18 },
        paragraph: { spacing: { before: 40, after: 40 } }
      },
      {
        id: "DocFooter",
        name: "Doc Footer",
        basedOn: "DocBody",
        run: { size: 16, color: deep },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 } }
      }
    ]
  };
}

// src/docx/primitives.ts
import {
  AlignmentType as AlignmentType2,
  BorderStyle as BorderStyle2,
  Footer,
  ImageRun,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType
} from "docx";
var HAIRLINE = "CCCCCC";
var hairline = { style: BorderStyle2.SINGLE, size: 4, color: HAIRLINE };
var borders = { top: hairline, bottom: hairline, left: hairline, right: hairline };
var noBorder = { style: BorderStyle2.NONE, size: 0, color: "FFFFFF" };
var noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
function pageA4() {
  return {
    page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
    }
  };
}
function logoRun(asset, bytes, widthPx) {
  const heightPx = Math.round(widthPx / (asset.widthPx / asset.heightPx));
  const type = asset.mime === "image/svg+xml" ? "svg" : asset.mime === "image/jpeg" ? "jpg" : "png";
  if (type === "svg") {
    return new ImageRun({ type: "png", data: bytes, transformation: { width: widthPx, height: heightPx } });
  }
  return new ImageRun({ type, data: bytes, transformation: { width: widthPx, height: heightPx } });
}
function masthead(kit, opts) {
  const left = [new Paragraph({ style: "DocTitle", children: [new TextRun(opts.title)] })];
  if (opts.subtitle) left.push(new Paragraph({ style: "DocSubtitle", children: [new TextRun(opts.subtitle)] }));
  const right = [];
  if (opts.logoBytes && kit.logos.light) {
    right.push(
      new Paragraph({
        alignment: AlignmentType2.RIGHT,
        children: [logoRun(kit.logos.light, opts.logoBytes, opts.logoWidthPx ?? 180)]
      })
    );
  }
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({ width: { size: 68, type: WidthType.PERCENTAGE }, borders: noBorders, verticalAlign: VerticalAlign.CENTER, children: left }),
          new TableCell({ width: { size: 32, type: WidthType.PERCENTAGE }, borders: noBorders, verticalAlign: VerticalAlign.CENTER, children: right.length ? right : [new Paragraph("")] })
        ]
      })
    ]
  });
}
function footerText(kit) {
  const parts = [kit.tenant.legalName];
  if (kit.legal.abn) parts.push(`ABN ${kit.legal.abn}`);
  if (kit.legal.address) parts.push(kit.legal.address);
  if (kit.legal.phone) parts.push(kit.legal.phone);
  if (kit.legal.web) parts.push(kit.legal.web);
  return parts.join("  |  ");
}
function footer(kit) {
  return new Footer({
    children: [
      new Paragraph({
        style: "DocFooter",
        border: { top: { style: BorderStyle2.SINGLE, size: 6, color: kit.palette.primary, space: 4 } },
        children: [new TextRun(footerText(kit))]
      })
    ]
  });
}
function h1(text) {
  return new Paragraph({ style: "DocH1", children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ style: "DocH2", children: [new TextRun(text)] });
}
function body(text) {
  return new Paragraph({ style: "DocBody", children: [new TextRun(text)] });
}
function small(text) {
  return new Paragraph({ style: "DocSmall", children: [new TextRun(text)] });
}
function spacer() {
  return new Paragraph({ style: "DocBody", children: [new TextRun("")] });
}
function headCell(kit, text, widthPct) {
  return new TableCell({
    width: widthPct ? { size: widthPct, type: WidthType.PERCENTAGE } : void 0,
    borders,
    shading: { type: ShadingType.CLEAR, fill: kit.palette.primary, color: "auto" },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ style: "DocTableHead", children: [new TextRun({ text, color: textOn(kit.palette.primary, kit.palette.ink) })] })]
  });
}
function bodyCell(kit, text, zebra, widthPct) {
  return new TableCell({
    width: widthPct ? { size: widthPct, type: WidthType.PERCENTAGE } : void 0,
    borders,
    shading: zebra ? { type: ShadingType.CLEAR, fill: kit.palette.ice, color: "auto" } : void 0,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ style: "DocTableCell", children: [new TextRun(text)] })]
  });
}
function dataTable(kit, opts) {
  const n = opts.head.length;
  const widths = opts.widths ?? opts.head.map(() => Math.floor(100 / n));
  const zebra = opts.zebra ?? true;
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: opts.head.map((h, i) => headCell(kit, h, widths[i])) }),
      ...opts.rows.map(
        (r, ri) => new TableRow({ cantSplit: true, children: r.map((c, ci) => bodyCell(kit, c, zebra && ri % 2 === 1, widths[ci])) })
      )
    ]
  });
}
function kvTable(kit, pairs, labelWidthPct = 30) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: pairs.map(
      ([k, v]) => new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: labelWidthPct, type: WidthType.PERCENTAGE },
            borders,
            shading: { type: ShadingType.CLEAR, fill: kit.palette.ice, color: "auto" },
            children: [new Paragraph({ style: "DocTableCell", children: [new TextRun({ text: k, bold: true })] })]
          }),
          new TableCell({
            width: { size: 100 - labelWidthPct, type: WidthType.PERCENTAGE },
            borders,
            children: [new Paragraph({ style: "DocTableCell", children: [new TextRun(v)] })]
          })
        ]
      })
    )
  });
}

// src/docx/pack.ts
import { Document, Packer } from "docx";
function createDocument(kit, opts) {
  const kitFooter = footer(kit);
  return new Document({
    creator: opts.creator ?? kit.tenant.displayName,
    title: opts.title,
    description: opts.description,
    styles: docxStyles(kit),
    sections: opts.sections.map((s) => ({
      ...s,
      properties: s.properties ?? pageA4(),
      footers: s.footers ?? { default: kitFooter }
    }))
  });
}
function toBlob(doc) {
  return Packer.toBlob(doc);
}
function toBuffer(doc) {
  return Packer.toBuffer(doc);
}
function toUint8Array(doc) {
  return Packer.toBuffer(doc).then((b) => new Uint8Array(b));
}

// src/xlsx/theme.ts
var theme_exports = {};
__export(theme_exports, {
  applyBodyRows: () => applyBodyRows,
  applyHeaderRow: () => applyHeaderRow,
  xlsxMasthead: () => xlsxMasthead,
  xlsxTheme: () => xlsxTheme
});
var argb = (hex) => `FF${hex}`;
function xlsxTheme(kit) {
  const { primary, ice, ink } = kit.palette;
  const thin = { style: "thin", color: { argb: argb("CCCCCC") } };
  return {
    headerFill: { type: "pattern", pattern: "solid", fgColor: { argb: argb(primary) } },
    headerFont: { name: kit.fonts.docBody, bold: true, color: { argb: argb(textOn(primary, ink)) }, size: 10 },
    zebraFill: { type: "pattern", pattern: "solid", fgColor: { argb: argb(ice) } },
    bodyFont: { name: kit.fonts.docBody, color: { argb: argb(ink) }, size: 10 },
    titleFont: { name: kit.fonts.heading, bold: true, color: { argb: argb(primary) }, size: 16 },
    border: { top: thin, bottom: thin, left: thin, right: thin }
  };
}
function applyHeaderRow(row, theme) {
  row.eachCell((cell) => {
    cell.fill = theme.headerFill;
    cell.font = theme.headerFont;
    cell.border = theme.border;
    cell.alignment = { vertical: "middle" };
  });
}
function applyBodyRows(ws, theme, from, to, columns) {
  for (let r = from; r <= to; r++) {
    const row = ws.getRow(r);
    const zebra = (r - from) % 2 === 1;
    for (let c = 1; c <= columns; c++) {
      const cell = row.getCell(c);
      cell.font = theme.bodyFont;
      cell.border = theme.border;
      if (zebra) cell.fill = theme.zebraFill;
    }
  }
}
function xlsxMasthead(wb, ws, kit, opts) {
  const theme = xlsxTheme(kit);
  const span = opts.span ?? 6;
  ws.mergeCells(1, 1, 1, Math.max(1, span - 2));
  const title = ws.getCell(1, 1);
  title.value = opts.title;
  title.font = theme.titleFont;
  ws.getRow(1).height = 30;
  let next = 2;
  if (opts.subtitle) {
    ws.mergeCells(2, 1, 2, Math.max(1, span - 2));
    const st = ws.getCell(2, 1);
    st.value = opts.subtitle;
    st.font = { ...theme.bodyFont, color: { argb: argb(kit.palette.deep) } };
    next = 3;
  }
  if (opts.logoBytes && kit.logos.light) {
    const asset = kit.logos.light;
    const width = opts.logoWidthPx ?? 160;
    const height = Math.round(width / (asset.widthPx / asset.heightPx));
    const ext = asset.mime === "image/jpeg" ? "jpeg" : "png";
    const id = wb.addImage({ buffer: opts.logoBytes, extension: ext });
    ws.addImage(id, { tl: { col: span - 2, row: 0 }, ext: { width, height } });
    const rowsNeeded = Math.ceil(height / 20) + 1;
    next = Math.max(next, rowsNeeded + 1);
  }
  return next + 1;
}

// src/template/fillDocx.ts
var fillDocx_exports = {};
__export(fillDocx_exports, {
  escapeXml: () => escapeXml,
  fillDocx: () => fillDocx,
  listSdtTags: () => listSdtTags,
  replaceSdtByTag: () => replaceSdtByTag
});
import JSZip from "jszip";
function escapeXml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function runXml(text, rPr) {
  const lines = text.split(/\r?\n/);
  const inner = lines.map((l) => `<w:t xml:space="preserve">${escapeXml(l)}</w:t>`).join("<w:br/>");
  return `<w:r>${rPr}${inner}</w:r>`;
}
function replaceSdtByTag(xml, tag, text) {
  let count = 0;
  const re = /<w:sdt>([\s\S]*?)<\/w:sdt>/g;
  const out = xml.replace(re, (whole, inner) => {
    const tagMatch = inner.match(/<w:tag\s+w:val="([^"]*)"/);
    if (!tagMatch || tagMatch[1] !== tag) return whole;
    const contentMatch = inner.match(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/);
    if (!contentMatch) return whole;
    const content = contentMatch[1];
    const rPr = (content.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) ?? [""])[0];
    const pPr = (content.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) ?? [""])[0];
    const isBlock = /<w:p[\s>]/.test(content);
    const replacement = isBlock ? `<w:p>${pPr}${runXml(text, rPr)}</w:p>` : runXml(text, rPr);
    count++;
    return whole.replace(contentMatch[0], `<w:sdtContent>${replacement}</w:sdtContent>`);
  });
  return { xml: out, count };
}
function listSdtTags(xml) {
  const tags = [];
  const re = /<w:tag\s+w:val="([^"]*)"/g;
  let m;
  while (m = re.exec(xml)) tags.push(m[1]);
  return tags;
}
async function fillDocx(template, bindings) {
  const zip = await JSZip.loadAsync(template);
  const parts = Object.keys(zip.files).filter((p) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(p));
  const filled = /* @__PURE__ */ new Set();
  for (const part of parts) {
    let xml = await zip.file(part).async("string");
    for (const [tag, value] of Object.entries(bindings)) {
      const r = replaceSdtByTag(xml, tag, value ?? "");
      if (r.count > 0) filled.add(tag);
      xml = r.xml;
    }
    zip.file(part, xml);
  }
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const unmatched = Object.keys(bindings).filter((t) => !filled.has(t));
  return { bytes, filled: [...filled], unmatched };
}

// src/preflight/check.ts
var check_exports = {};
__export(check_exports, {
  allowedHex: () => allowedHex,
  extractFacts: () => extractFacts,
  preflight: () => preflight,
  preflightDocx: () => preflightDocx
});
import JSZip2 from "jszip";
function allowedHex(kit) {
  const s = /* @__PURE__ */ new Set([
    kit.palette.primary,
    kit.palette.deep,
    kit.palette.ice,
    kit.palette.ink,
    "FFFFFF",
    HAIRLINE,
    mutedInk(kit),
    "AUTO"
  ]);
  if (kit.palette.accent) s.add(kit.palette.accent);
  return s;
}
function preflight(kit, facts) {
  const checks = [];
  const kitLogoUrls = [kit.logos.light?.url, kit.logos.dark?.url, kit.logos.mark?.url].filter(Boolean);
  const badLogos = (facts.logoSources ?? []).filter((u) => !kitLogoUrls.includes(u));
  checks.push({ id: "logo", pass: badLogos.length === 0, detail: badLogos.length ? `logo not from kit: ${badLogos.join(", ")}` : void 0 });
  const expected = kit.logos.light ? kit.logos.light.widthPx / kit.logos.light.heightPx : void 0;
  const badRatio = (facts.logoPlacements ?? []).filter((p) => expected !== void 0 && Math.abs(p.width / p.height - expected) / expected > 0.02);
  checks.push({ id: "ratio", pass: badRatio.length === 0, detail: badRatio.length ? `logo stretched: ${badRatio.map((p) => `${p.width}\xD7${p.height}`).join(", ")}` : void 0 });
  const allowed = allowedHex(kit);
  const foreign = [...new Set(facts.usedHex.map((h) => h.toUpperCase()))].filter((h) => !allowed.has(h));
  const headContrast = contrastRatio("FFFFFF", kit.palette.primary);
  const paletteDetail = foreign.length ? `colours outside kit: ${foreign.join(", ")}` : headContrast < 3 ? `white on primary is ${headContrast.toFixed(1)}:1` : void 0;
  checks.push({ id: "palette", pass: foreign.length === 0 && headContrast >= 3, detail: paletteDetail });
  const kitFonts = /* @__PURE__ */ new Set([kit.fonts.heading, kit.fonts.body, kit.fonts.docBody]);
  const badFonts = [...new Set(facts.usedFonts)].filter((f) => !kitFonts.has(f));
  checks.push({ id: "fonts", pass: badFonts.length === 0, detail: badFonts.length ? `fonts outside kit: ${badFonts.join(", ")}` : void 0 });
  checks.push({ id: "flat", pass: !facts.hasEffects, detail: facts.hasEffects ? "gradient/shadow/effect present" : void 0 });
  const footerOk = facts.hasFooter && (facts.footerText ?? "").includes(kit.tenant.legalName);
  checks.push({ id: "footer", pass: footerOk, detail: !facts.hasFooter ? "no footer" : footerOk ? void 0 : `footer does not carry "${kit.tenant.legalName}"` });
  const ok = checks.every((c) => c.pass);
  const line = "Brand check: " + checks.map((c) => c.pass ? `\u2713 ${c.id}` : `\u2717 ${c.id} \u2192 ${c.detail ?? "failed"}`).join(" ");
  return { ok, checks, line };
}
var HEX_ATTR = /w:(?:color|fill|themeColor)="([0-9A-Fa-f]{6}|auto)"/g;
var FONT_ATTR = /w:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"/g;
var EFFECT_TAGS = /<w:(?:shadow|glow|reflection|effect|gradFill|outline|emboss|imprint)[\s/>]|<a:gradFill|<a:effectLst>\s*<a:/;
var EXTENT = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/g;
async function extractFacts(docx) {
  const zip = await JSZip2.loadAsync(docx);
  const parts = Object.keys(zip.files).filter((p) => /^word\/(document|styles|header\d*|footer\d*|numbering)\.xml$/.test(p));
  const usedHex = /* @__PURE__ */ new Set();
  const usedFonts = /* @__PURE__ */ new Set();
  let hasEffects = false;
  let hasFooter = false;
  let footerText2 = "";
  const logoPlacements = [];
  for (const part of parts) {
    const xml = await zip.file(part).async("string");
    for (const m of xml.matchAll(HEX_ATTR)) usedHex.add(m[1].toUpperCase());
    for (const m of xml.matchAll(FONT_ATTR)) usedFonts.add(m[1]);
    if (EFFECT_TAGS.test(xml)) hasEffects = true;
    for (const m of xml.matchAll(EXTENT)) logoPlacements.push({ width: Number(m[1]), height: Number(m[2]) });
    if (/^word\/footer\d*\.xml$/.test(part)) {
      hasFooter = true;
      footerText2 += xml.replace(/<[^>]+>/g, " ");
    }
  }
  return { usedHex: [...usedHex], usedFonts: [...usedFonts], hasEffects, hasFooter, footerText: footerText2.replace(/\s+/g, " ").trim(), logoPlacements };
}
async function preflightDocx(kit, docx) {
  return preflight(kit, await extractFacts(docx));
}

// src/index.ts
import { NEUTRAL_BRAND_KIT as NEUTRAL_BRAND_KIT2, BRAND_KIT_VERSION as BRAND_KIT_VERSION2, validateTenantBrandKit, toHex6 as toHex62, isHex6 } from "@eq-solutions/contracts";
export {
  BRAND_KIT_VERSION2 as BRAND_KIT_VERSION,
  NEUTRAL_BRAND_KIT2 as NEUTRAL_BRAND_KIT,
  brand_exports as brand,
  docx_exports as docx,
  isHex6,
  check_exports as preflight,
  fillDocx_exports as template,
  toHex62 as toHex6,
  validateTenantBrandKit,
  theme_exports as xlsx
};
//# sourceMappingURL=documents.esm.js.map
