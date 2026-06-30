// Edge function: parse RFP PDF and build a structured Excel workbook
// - Uses unpdf for text extraction + image extraction (works in Deno/edge runtime)
// - Uses exceljs for .xlsx generation
// - Heuristic structure detection: top-level numbered headings -> sheets,
//   subheadings stay in the same sheet, tables get a separate sheet (one per
//   table occurrence), images extracted from the PDF are embedded inline in
//   the content sheet at their approximate position.

import ExcelJS from "npm:exceljs@4.4.0";
import { createClient } from "npm:@supabase/supabase-js";
import { extractText, getDocumentProxy, extractImages } from "npm:unpdf@1.6.2";
import { PNG } from "npm:pngjs@6.0.0";
import { verifyToken } from "../shared.ts";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const supabaseAdmin = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false },
});

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-app-token, x-session-id",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

type Block =
  | {
      type: "heading" | "subheading" | "paragraph";
      text: string;
      level?: number;
      page: number;
    }
  | {
      type: "table-row";
      text: string;
      cells: string[];
      page: number;
    }
  | {
      type: "image";
      page: number;
      imageId: number; // index into PdfImage[]
    };

type Section = {
  number: string;
  title: string;
  page: number;
  blocks: Block[];
};

type PdfImage = {
  pngBytes: Uint8Array;
  width: number;
  height: number;
  page: number;
};

// --- PDF text + lightweight structure detection -------------------------------

async function extractStructured(bytes: Uint8Array): Promise<{
  pages: {
    pageNumber: number;
    lines: string[];
  }[];
}> {
  const pdf = await getDocumentProxy(bytes);
  const { text } = await extractText(pdf, { mergePages: false });
  const pageTexts = Array.isArray(text) ? text : [text];
  const pages = pageTexts.map((p, idx) => {
  const allLines = p
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(Boolean);

// Ignore first 3 lines (header)
const HEADER_LINES = 2;
const FOOTER_LINES = 2;
const bodyWithoutHeader =
    allLines.slice(HEADER_LINES);

const body =
    bodyWithoutHeader.slice(
        0,
        Math.max(
            0,
            bodyWithoutHeader.length - FOOTER_LINES
        )
    );

const raw = body.filter(line => {

    if (
        /^page\s+\d+/i.test(line) ||
        /^\d+$/.test(line) ||
        /^\d+\s*\/\s*\d+$/.test(line)
    ) {
        return false;
    }

    return true;
});
    const out: string[] = [];
    let prevBlank = true;
    for (const l of raw) {
      if (l.length === 0) {
        if (!prevBlank) out.push("");
        prevBlank = true;
      } else {
        out.push(l);
        prevBlank = false;
      }
    }
    return {
  pageNumber: idx + 1,
  lines: out,
};
  });
  return { pages };
}

// Encode raw pixel data (from unpdf extractImages) into PNG bytes.
function encodeToPng(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  channels: 1 | 3 | 4,
): Uint8Array {
  const png = new PNG({ width, height });
  // png.data is RGBA, length = width*height*4
  const out = png.data;
  const px = width * height;
  if (channels === 4) {
    out.set(data);
  } else if (channels === 3) {
    for (let i = 0; i < px; i++) {
      out[i * 4 + 0] = data[i * 3 + 0];
      out[i * 4 + 1] = data[i * 3 + 1];
      out[i * 4 + 2] = data[i * 3 + 2];
      out[i * 4 + 3] = 255;
    }
  } else {
    // grayscale
    for (let i = 0; i < px; i++) {
      const v = data[i];
      out[i * 4 + 0] = v;
      out[i * 4 + 1] = v;
      out[i * 4 + 2] = v;
      out[i * 4 + 3] = 255;
    }
  }
  const buf = PNG.sync.write(png);
  return new Uint8Array(buf);
}

async function extractAllImages(bytes: Uint8Array, pageCount: number): Promise<PdfImage[]> {
  const pdf = await getDocumentProxy(bytes);
  const images: PdfImage[] = [];
  for (let p = 1; p <= pageCount; p++) {
    try {
      const list = await extractImages(pdf, p);
      for (const img of list) {
        // Skip very tiny images (decorative bullets, lines)
        if (img.width < 32 || img.height < 32) continue;
        try {
          const pngBytes = encodeToPng(img.data, img.width, img.height, img.channels);
          images.push({ pngBytes, width: img.width, height: img.height, page: p });
        } catch (e) {
          console.warn(`Image encode failed page ${p}:`, e instanceof Error ? e.message : e);
        }
      }
    } catch (e) {
      console.warn(`Image extract failed page ${p}:`, e instanceof Error ? e.message : e);
    }
  }
  return images;
}

// Merge consecutive paragraph blocks into single paragraph blocks per section.
function mergeParagraphs(blocks: Block[]): Block[] {
  const out: Block[] = [];
  let buffer: string[] = [];
  let currentPage = 0;

  const flush = () => {
    if (buffer.length > 0) {
      const text = buffer.join(" ").trim();
      if (text.length > 30) {
        out.push({
          type: "paragraph",
          text,
          page: currentPage,
        });
      }
      buffer = [];
    }
  };

  for (const b of blocks) {
    if (b.type !== "paragraph") {
      flush();
      out.push(b);
      continue;
    }

    const line = b.text.trim();

    // 🔹 Ignore artificial breaks
    if (line === "\u0000BREAK") {
      flush();
      continue;
    }

    if (buffer.length === 0) {
      buffer.push(line);
      currentPage = b.page;
      continue;
    }

    const prev = buffer[buffer.length - 1];

    const isNewParagraph =
      /[.!?:]$/.test(prev) &&   // previous sentence ended
      line.length > 40;        // next line is meaningful

    if (isNewParagraph) {
      flush();
      buffer.push(line);
      currentPage = b.page;
    } else {
      buffer.push(line);
    }
  }

  flush();
  return out;
}

const TOP_HEADING_RE =
  /^(?:section|chapter|part)\s+(\d{1,3})[\.\):\-—\s]+(.{2,120}?)$|^(\d{1,3})[\.\)]?\s+([A-Z][A-Za-z0-9 &/,\-—:]{2,120})$|^([A-Z][A-Z\s/&\-]{5,120})$/;
const SUB_HEADING_RE = /^(\d{1,2}(?:\.\d{1,2}){1,3})\s+([A-Z].{1,100}?)$/;

function splitMultiGap(line: string): string[] | null {

  if (line.includes("\t")) {
    const parts = line.split(/\t+/).map((s) => s.trim());
    return parts.length >= 2 ? parts : null;
  }

  const parts = line
    .split(/\s{3,}|\t+| {2,}(?=[A-Z0-9])/)
    .map((s) => s.trim())
    .filter(Boolean);

  if (parts.length < 2 || parts.length > 20)
    return null;

  if (
  line.length > 500 &&
  !/\s{3,}/.test(line)
)
  return null;

  // ── GENERIC 2-COLUMN RULE (handles Position|Qualifications type tables) ──
  // If exactly 2 parts and the first is a short label/name (≤45 chars, no
  // leading number like "1."), treat as valid 2-cell row regardless of
  // second cell length.  This catches "Project Manager   1. Education:..."
  if (
  parts.length === 2 &&
  parts[0].length <= 35 &&
  parts[0].split(/\s+/).length <= 4 &&
  !/^\d+[\.\)]/.test(parts[0]) &&
  (
    /^[A-Z]/.test(parts[0]) ||
    /:$/.test(parts[0])
  ) &&
  parts[1].length >= 8
) {
  return parts;
}
  // ─────────────────────────────────────────────────────────────────────────
// detect title + long description tables
const twoCol =
  line.match(
    /^([A-Za-z0-9 &\/\-,()]{2,50}?)\s{2,}(.{10,})$/
  );

if (
  twoCol &&
  twoCol[1].split(/\s+/).length <= 5 &&
  twoCol[2].length >= 10 &&
  !/[.!?]$/.test(twoCol[1])
) {
  return [
    twoCol[1].trim(),
    twoCol[2].trim(),
  ];
}
  // For 3+ parts: reject sentence fragments
  if (/[.!?]$/.test(line) && parts.length <= 2)
  return null;
  const anyLong = parts.some((p) => p.split(/\s+/).length > 8 || p.length > 50);
  if (anyLong && parts.length <= 2)
  return null;

  return parts;
}
function splitKeyedRow(line: string): string[] | null {
  const m = line.match(
    /^(PQ\s*\d+(?:\.\d+)?|[A-Z]{1,3}\d+(?:\.\d+)?|\d+(?:\.\d+)?)\s+(.+)$/i
  );

  if (!m) return null;

  const key = m[1].trim();
  const rest = m[2].trim();

  let parts = rest
  .split(/\s{3,}|\t+| {2,}(?=[A-Z0-9])/)
  .map((s) => s.trim());

// fallback for collapsed PDF spacing
if (parts.length < 2) {

  // try detecting sentence boundaries
  const tokens = rest.split(/\s+/);

  if (tokens.length >= 6) {

    parts = [];

    let current = "";

    for (const t of tokens) {

      current += (current ? " " : "") + t;

      // split at strong boundary words
      if (
        /^(shall|should|must|will|required|provide|submit)$/i.test(t)
      ) {
        parts.push(current.trim());
        current = "";
      }
    }

    if (current.trim()) {
      parts.push(current.trim());
    }
  }
}

  if (parts.length >= 2) {
    return [key, ...parts];
  }

  return [key, rest];
}
// Detect a numbered table row like "1. Project Manager Nos. 1" or "1 Project Manager Nos. 1"
function splitNumberedRow(line: string): string[] | null {
  const m = line.match(/^(\d{1,3})[\.\)]?\s+(.+)$/);
  if (!m) return null;

  const serial = m[1];
  const rest = m[2].trim();

  if (rest.length > 200) return null;

  // ── NEW GUARDS: reject qualification-style sentences ──────────────────────
  // Pattern: "Education: ...", "Total Experience: ...", "Should have..."
  if (/^[A-Za-z][A-Za-z ]{2,}:\s/.test(rest)) return null;
  // Long sentence-like content (>60 chars AND >7 words)
  if (rest.length > 60 && rest.split(/\s+/).length > 7) return null;
  // Complete sentence ending with punctuation
  if (/[.!?]$/.test(rest) && rest.length > 30) return null;
  // ─────────────────────────────────────────────────────────────────────────
  // Try structured spacing first
  const parts = rest
  .split(/\s{3,}|\t+| {2,}(?=[A-Z0-9])/)
  .map((s) => s.trim());

  if (parts.length >= 2) {
    return [serial, ...parts];
  }

  // Fallback for collapsed PDF spacing
  const words = rest.split(/\s+/);
  const looksLikeTable =
  /\b(?:nos?|qty|quantity|uom|rate|amount|price|cost)\b/i.test(rest);

  if (
  !looksLikeTable &&
  words.length > 12
) {
  return null;
}

  // Detect patterns like:
// "Project Manager Nos. 1"
// "Helpdesk Executive Nos. 6"

// Generic structured row detection
// Example:
// "Project Manager Nos. 1"
// "Database Admin Unit 2"
// "Fiber Cable Km 12"

const wordParts = rest.trim().split(/\s+/);

if (
  wordParts.length >= 3 &&
  /^\d+(\.\d+)?$/.test(wordParts[wordParts.length - 1])
) {

  const quantity =
    wordParts[wordParts.length - 1];

  const possibleUnit =
    wordParts[wordParts.length - 2];

  if (
    possibleUnit.length <= 12 &&
    /^[A-Za-z%]+\.?$/.test(possibleUnit)
  ) {

    let description =
      wordParts.slice(0, -2).join(" ");
      if (
  description
    .toLowerCase()
    .endsWith(
      possibleUnit.toLowerCase()
    )
) {
  description =
    description.slice(
      0,
      -possibleUnit.length
    ).trim();
}

    return [
      serial,
      description,
      possibleUnit.replace(/\./g, ""),
      quantity,
    ];
  }
}

// Generic fallback
if (
  words.length >= 4 &&
  words.length <= 12 &&
  !/[.!?]$/.test(rest)
) {
  return [
    serial,
    words.slice(0, -2).join(" "),
    words[words.length - 2],
    words[words.length - 1],
  ];
}

  return [serial, rest];
}
function mergeWrappedHeaders(rows: string[][]): string[][] {

  if (rows.length < 2) return rows;

  const first = rows[0];
  const second = rows[1];

  const merged =
    first.map((c, i) => {
      const next = second[i] ?? "";

      // merge short wrapped header fragments
  if (
  c.length < 20 &&
  next.length < 20 &&
  !/^\d/.test(next) &&
  /^[A-Za-z]/.test(c) &&
  /^[A-Za-z]/.test(next)
) {
        return `${c} ${next}`.trim();
      }

      return c;
    });

  return [merged, ...rows.slice(2)];
}
function looksLikeDelimitedRow(
  line: string
): boolean {

  return (
    /\|/.test(line) ||
    /:{1,}/.test(line) ||
    /\b(?:qty|uom|amount|price|rate)\b/i.test(line)
  );
}
function looksLikeStructuredRow(line: string): boolean {

  const parts =
    line.trim().split(/\s+/);

  if (parts.length < 3)
    return false;

  const numericCount =
    parts.filter(p =>
      /^\d+(\.\d+)?$/.test(p)
    ).length;

  const shortCount =
    parts.filter(p =>
      p.length <= 12
    ).length;

  return (
  numericCount >= 1 &&
  shortCount >= 3 &&
  parts.length <= 12 &&
  !/[.!?]$/.test(line)
);
}
function reconstructCells(line: string): string[] | null {

  const tokens = line.trim().split(/\s+/);

  if (tokens.length < 3)
    return null;

  // detect trailing numeric quantity/value
  const last = tokens[tokens.length - 1];

  if (!/^\d+(\.\d+)?$/.test(last))
    return null;

  // detect short unit-like token before quantity
  const prev = tokens[tokens.length - 2];

  if (
    prev.length <= 12 &&
    /^[A-Za-z%\.]+$/.test(prev)
  ) {

    const description =
      tokens.slice(0, -2).join(" ");

    return [
  tokens.slice(0, -2).join(" ").trim(),
  prev.replace(/\./g, "").trim(),
  last.trim(),
];
  }

  // fallback:
  // first token separate, rest grouped
  return [
    tokens[0],
    tokens.slice(1).join(" "),
  ];
}
// After tokenizing pages, detect *table regions*: clusters of consecutive
// numbered-row lines (>=2 in a row). Mark them with a shared tableId.
type PendingTable = {
  rows: string[];
  page: number;
};
function flushPendingTable(
  pending: PendingTable | null,
  current: Section | null
) {

  if (!pending || !current)
    return;
const parsedRows =
  pending.rows
    .map(line =>
      line
        .split(" || ")
        .map(x => x.trim())
    )
    .filter(r => r.length >= 2);

if (
  parsedRows.length < 2 &&
  !parsedRows.some(r => r.length >= 3)
) {
  return;
}

// dominant column count
const dominant =
  parsedRows
    .map(r => r.length)
    .sort(
      (a, b) =>
        parsedRows.filter(x => x.length === b).length -
        parsedRows.filter(x => x.length === a).length
    )[0];

const normalizedRows =
  parsedRows.map(r => {

    const copy = [...r];

    while (copy.length < dominant)
      copy.push("");

    // merge overflow into last column
    if (copy.length > dominant) {

      const extra =
        copy.slice(dominant - 1).join(" ");

      copy.splice(
        dominant - 1,
        copy.length,
        extra
      );
    }

    return copy;
  });

// merge continuation rows
for (let i = 1; i < normalizedRows.length; i++) {

  const currentRow = normalizedRows[i];
  const prevRow = normalizedRows[i - 1];

  if (
    currentRow.length <= 2 &&
    prevRow.length >= 3 &&
    currentRow.join(" ").length < 80 &&
    !/^\d+[\.\)]?$/.test(currentRow[0] ?? "")
  ) {

    prevRow[prevRow.length - 1] +=
      "\n" + currentRow.join(" ");

    normalizedRows.splice(i, 1);

    i--;
  }
}
  normalizedRows.forEach((cells, idx) => {

    while (cells.length < dominant)
      cells.push("");

    current.blocks.push({
      type: "table-row",
      text: pending.rows[idx],
      cells,
      page: pending.page,
    });
  });
}
function detectTableRow(
  line: string
): string[] | null {

  const candidates = [

    splitMultiGap(line),
    splitKeyedRow(line),
    splitNumberedRow(line),

  ].filter(Boolean) as string[][];

  if (candidates.length === 0)
    return null;


  // choose best candidate
  const best =
    candidates.sort(
      (a, b) => b.length - a.length
    )[0];

  // reject paragraph-like rows
  const joined =
    best.join(" ");

  // too sentence-like
  if (
    joined.length > 120 &&
    /[.!?]$/.test(joined)
  ) {
    return null;
  }

  // must contain structure
  const structuredCount =
    best.filter(cell =>
      cell.length > 0 &&
      cell.length < 80
    ).length;

  if (structuredCount < 2)
    return null;

  return best;
}
function buildSections(
  pages: {
    pageNumber: number;
    lines: string[];
  }[]
): Section[] {
  const sections: Section[] = [];
  let current: Section | null = null;
  let pendingTable: PendingTable | null = null;
  const pushIntro = (page: number) => {
    if (!current) {
      current = { number: "0", title: "Preamble", page, blocks: [] };
      sections.push(current);
    }
  };
  pages.forEach((p) => {
  const lines = p.lines;
  const page = p.pageNumber;
  // keep table alive across pages
    for (const raw of lines) {
      if (raw.length === 0) {
        if (current && current.blocks.length > 0) {
          const last = current.blocks[current.blocks.length - 1];
          if (last.type === "paragraph" && last.text !== "\u0000BREAK") {
            current.blocks.push({ type: "paragraph", text: "\u0000BREAK", page });
          }
        }
        // blank lines should not aggressively break tables
continue;
      }
      const line = raw.replace(/\s+/g, " ").trim();
      if (
    /^page\s+\d+/i.test(line) ||
    /^\d+\s*of\s*\d+$/i.test(line) ||
    /^\d+\s*\/\s*\d+$/.test(line) ||
    /^confidential$/i.test(line) ||
    /^copyright/i.test(line)
) {
    continue;
}
      if (!line) continue;

      const top = line.match(TOP_HEADING_RE);
      if (top) {
        flushPendingTable(
  pendingTable,
  current
);

pendingTable = null;
        current = {
          number: (top[1] ?? top[3]) as string,
          title: ((top[2] ?? top[4]) as string).trim(),
          page,
          blocks: [],
        };
        sections.push(current);
        continue;
      }
      const sub = line.match(SUB_HEADING_RE);
      if (sub) {
        flushPendingTable(
  pendingTable,
  current
);

pendingTable = null;
  pushIntro(page);
        current!.blocks.push({
          type: "subheading",
          text: `${sub[1]} ${sub[2].trim()}`,
          level: sub[1].split(".").length,
          page,
          
        });
        continue;
      }
      const parsedTableRow =
  detectTableRow(line);

const looksTableLike =
  parsedTableRow !== null;

if (looksTableLike) {

  if (!pendingTable) {
    pendingTable = {
      rows: [],
      page,
    };
  }

  pendingTable.rows.push(
  parsedTableRow.join(" || ")
);

  continue;
}
// Close table only for strong non-table paragraph blocks
  const looksLikeLongParagraph =
  line.length > 220 &&
  line.split(/\s+/).length > 35 &&
  /[.!?]$/.test(line);

if (looksLikeLongParagraph) {
  flushPendingTable(
  pendingTable,
  current
);

pendingTable = null;
  pushIntro(page);
  current!.blocks.push({
    type: "paragraph",
    text: line,
    page,
  });

  continue;
}
if (pendingTable) {

  flushPendingTable(
    pendingTable,
    current
  );

  pendingTable = null;
}
      pushIntro(page);
      current!.blocks.push({ type: "paragraph", text: line, page });
    }
  });
flushPendingTable(
  pendingTable,
  current
);

pendingTable = null;
for (const s of sections) {

  s.blocks =
    mergeParagraphs(s.blocks).filter(
      (b) =>
        !(
          b.type === "paragraph" &&
          b.text === "\u0000BREAK"
        ),
    );
}

return sections;
}

// Attach images to sections by page overlap. An image on page P is attached
// to the section whose page range contains P. We insert image blocks at the
// position of the first block on that page (so order ~ matches reading order).
function attachImagesToSections(sections: Section[], images: PdfImage[]) {
  if (images.length === 0 || sections.length === 0) return;

  // For each image, find target section
  for (let imgIdx = 0; imgIdx < images.length; imgIdx++) {
    const img = images[imgIdx];
    let target: Section | null = null;
    for (let i = 0; i < sections.length; i++) {
      const s = sections[i];
      const next = sections[i + 1];
      const startPage = s.page;
      const endPage = next ? next.page : Number.MAX_SAFE_INTEGER;
      if (img.page >= startPage && img.page < endPage) {
        target = s;
        break;
      }
      if (img.page < startPage) {
        target = s;
        break;
      }
    }
    if (!target) target = sections[sections.length - 1];

    // Find first block on the same page or after, insert before it
    let insertIdx = target.blocks.length;
    for (let i = 0; i < target.blocks.length; i++) {
      const b = target.blocks[i];
      const bp = "page" in b ? b.page : 0;
      if (bp >= img.page) { insertIdx = i; break; }
    }
    target.blocks.splice(insertIdx, 0, { type: "image", page: img.page, imageId: imgIdx });
  }
}

// --- Excel builder ------------------------------------------------------------

function safeSheetName(name: string, used: Set<string>): string {
  let n = name.replace(/[:\\/?*\[\]]/g, " ").replace(/\s+/g, " ").trim();
  if (n.length > 28) n = n.slice(0, 28);
  if (!n) n = "Section";
  let candidate = n;
  let i = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` (${i++})`;
    candidate = n.slice(0, 28 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

const COMPLIANCE_HEADERS = ["Compliance (Yes/No)", "Remarks"];

function styleHeaderRow(row: ExcelJS.Row) {
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF002060" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FFBFBFBF" } },
      left: { style: "thin", color: { argb: "FFBFBFBF" } },
      bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
      right: { style: "thin", color: { argb: "FFBFBFBF" } },
    };
  });
  row.height = 24;
}

function styleSubheading(row: ExcelJS.Row) {
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FF1F4E78" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF1F8" } };
    cell.alignment = { vertical: "middle", wrapText: true };
  });
}

function styleTableMarker(row: ExcelJS.Row) {
  row.eachCell((cell) => {
    cell.font = { italic: true, bold: true, color: { argb: "FF1F4E78" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F7FC" } };
    cell.alignment = { vertical: "middle", wrapText: true };
  });
}
function applyBorders(row: ExcelJS.Row, cols: number) {
  for (let c = 1; c <= cols; c++) {
    row.getCell(c).border = {
      top: { style: "thin", color: { argb: "FFBFBFBF" } },
      bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
      left: { style: "thin", color: { argb: "FFBFBFBF" } },
      right: { style: "thin", color: { argb: "FFBFBFBF" } },
    };
  }
}
function buildContentSheet(
  wb: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  blocks: Block[],
  images: PdfImage[],
) {
  ws.columns = [
    { header: "S.No.", width: 8 },
    { header: "Technical Specifications", width: 70 },
    { header: COMPLIANCE_HEADERS[0], width: 22 },
    { header: COMPLIANCE_HEADERS[1], width: 36 },
  ];
  styleHeaderRow(ws.getRow(1));
  ws.getColumn(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  ws.getColumn(3).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  ws.getColumn(4).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  ws.views = [{ state: "frozen", ySplit: 1 }];

  const seenTableMarkers = new Set<string>();
  let sno = 1;

  for (const b of blocks) {
    if (b.type === "table-row") {

  const tableName =
    (b as any).__sheetName;

  if (!tableName)
    continue;

  if (!seenTableMarkers.has(tableName)) {

    seenTableMarkers.add(tableName);

    const r = ws.addRow([
      sno++,
      `→ See table on sheet: "${tableName}"`,
      "",
      ""
    ]);

    styleTableMarker(r);

    applyBorders(r, 4);
  }

  continue;
}

    if (b.type === "subheading") {
  const r = ws.addRow([sno++, b.text, "", ""]);

  applyBorders(r, 4);

  styleSubheading(r);

  r.getCell(1).alignment = {
    vertical: "middle",
    horizontal: "center",
  };

  r.getCell(3).alignment = {
    vertical: "middle",
    horizontal: "center",
  };

  continue;
    }

    if (b.type === "image") {
      const img = images[b.imageId];
      if (!img) continue;
      const maxW = 480;
      const maxH = 320;
      const ratio = Math.min(maxW / img.width, maxH / img.height, 1);
      const dispW = Math.max(64, Math.round(img.width * ratio));
      const dispH = Math.max(64, Math.round(img.height * ratio));

      const r = ws.addRow([sno++, "", "", ""]);
      applyBorders(r, 4);
      r.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
      r.getCell(3).alignment = { vertical: "middle", horizontal: "center" };
      r.height = Math.max(40, Math.round(dispH * 0.75) + 4);

      try {
        const imageId = wb.addImage({ buffer: img.pngBytes.buffer, extension: "png" });
        const rowIdx = r.number;
        ws.addImage(imageId, {
          tl: { col: 1.05, row: rowIdx - 1 + 0.05 } as any,
          ext: { width: dispW, height: dispH },
          editAs: "oneCell",
        });
      } catch (e) {
        r.getCell(2).value = `[Image from page ${img.page} — failed to embed]`;
        console.warn("addImage failed:", e instanceof Error ? e.message : e);
      }
      continue;
    }

    // paragraph
    const r = ws.addRow([sno++, b.text, "", ""]);
    r.alignment = { vertical: "top", wrapText: true };
    r.getCell(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    r.getCell(3).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    r.getCell(4).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    // Add borders to every paragraph row
    applyBorders(r, 4);
    r.height = Math.max(
  24,
  Math.min(
    Math.ceil(b.text.length / 90) * 18,
    140
  )
);
    }

  if (sno === 1) {
    const r = ws.addRow([1, "(No prose content detected for this section.)", "", "", ""]);
    r.getCell(1).alignment = { vertical: "top", horizontal: "center" };
  }
  // Note: No data validation dropdown on Compliance column — free text entry.
}
function buildTableSheet(ws: ExcelJS.Worksheet, tableRows: Block[]) {
  const tRows = tableRows.filter(
    (r): r is Extract<Block, { type: "table-row" }> => r.type === "table-row"
  );

  let rows = tRows.map(r => r.cells);
  rows = mergeWrappedHeaders(rows);
  const dominantCols =
  rows
    .map(r => r.length)
    .sort((a, b) =>
      rows.filter(x => x.length === b).length -
      rows.filter(x => x.length === a).length
    )[0];

rows = rows.filter(r => {

  // allow slight mismatch
  if (
    Math.abs(r.length - dominantCols) <= 1
  ) {
    return true;
  }

  // preserve long descriptive rows
  return (
    r.join(" ").length > 40
  );
});
  if (rows.length === 0) return;

  // --- HEADER --- 
  

// table headers (dynamic)
const maxCols = Math.max(
  3,
  ...rows.map(r => r.length)
);
// normalize rows to same column count
rows = rows.map(r => {

  const copy = [...r];

  while (copy.length < maxCols)
    copy.push("");

  // merge overflow into last column
  if (copy.length > maxCols) {

    const extra =
      copy.slice(maxCols - 1).join(" ");

    copy.splice(
      maxCols - 1,
      copy.length,
      extra
    );
  }

  return copy;
});
// detect header row
const firstRow = rows[0] || [];

const numericCount = rows.filter(r =>
  /^\d+[\.\)]?$/.test((r[0] ?? "").trim())
).length;

const firstColLooksNumeric =
  numericCount >= Math.ceil(rows.length * 0.7);

const isHeaderRow =

  firstRow.length >= 2 &&

  firstRow.filter(cell => {

    const clean =
      cell.trim();

    return (
      clean.length > 0 &&
      clean.length <= 80 &&
      /^[A-Za-z][A-Za-z0-9\s()./%:#,&\-]*$/
        .test(clean)
    );

  }).length >=
    Math.ceil(firstRow.length * 0.6) &&

  !firstRow.every(cell =>
    /^\d+(\.\d+)?$/.test(cell.trim())
  ) &&

  rows
    .slice(1, 4)
    .some(r =>
      r.some(cell =>
        /\d/.test(cell)
      )
    );

// use original headers if available
let headerRow = isHeaderRow ? firstRow : null;

// Smart fallback headers for manpower/resource tables
const dataRows = isHeaderRow ? rows.slice(1) : rows;
  if (firstColLooksNumeric) {
  ws.getCell(1, 1).value = "S.No.";
} else {
  ws.mergeCells(1, 1, 2, 1);
}

// dynamic headers
let effectiveHeaders: string[] = [];

if (headerRow) {
  effectiveHeaders = headerRow;
} else {
  // Generic fallback headers
  effectiveHeaders = Array.from(
  { length: maxCols },
  (_, i) =>
    i === 0 && firstColLooksNumeric
      ? "S.No."
      : ""
);
}

const techStartCol = 2;
const visibleHeaders =
  firstColLooksNumeric &&
  effectiveHeaders.length > 0 &&
  /^(s\.?no|sr\.?|serial)/i.test(effectiveHeaders[0])
    ? effectiveHeaders.slice(1)
    : effectiveHeaders;
const renderedCols = maxCols;

const techEndCol =
  techStartCol + renderedCols - 1;

// Dynamic merge: "Technical Specifications" spans exactly the data columns
const techMergeEnd = techEndCol;
if (techMergeEnd > techStartCol) {
  ws.mergeCells(1, techStartCol, 1, techMergeEnd);
} 
ws.getCell(1, techStartCol).value = "Technical Specifications";

// Compliance and Remarks go right after the last data column
ws.getCell(1, techMergeEnd + 1).value = "Compliance (Yes/No)";
ws.getCell(1, techMergeEnd + 2).value = "Remarks";

styleHeaderRow(ws.getRow(1));
const secondHeaderRow =
  ws.getRow(2);
const headerOffset = 2;
visibleHeaders.forEach((h, i) => {
  const cleanHeader =
  h
    ?.replace(/\s+/g, " ")
    .trim() || "";

  secondHeaderRow
    .getCell(i + headerOffset)
    .value = cleanHeader;
});
for (let c = 1; c <= ws.columnCount; c++) {

  const cell = secondHeaderRow.getCell(c);

  cell.font = { bold: true };

  cell.fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFD9EAF7" },
};

  cell.alignment = {
    vertical: "middle",
    horizontal: "center",
    wrapText: true,
  };

  cell.border = {
    top: { style: "thin" },
    bottom: { style: "thin" },
    left: { style: "thin" },
    right: { style: "thin" },
  };
}

secondHeaderRow.height = 28;
  // --- DATA ---
  let sno = 1;
dataRows.forEach((row, i) => {
const startRow =
  firstColLooksNumeric ? 3 : 2;
const excelRow = ws.getRow(i + startRow);

  let col = 1;

  // S.No (ONLY ONCE)
  if (firstColLooksNumeric) {
  excelRow.getCell(col++).value = sno++;
}
  
const actualRow = row;
  const normalized =
  firstColLooksNumeric &&
  /^\d+[\.\)]?$/.test((actualRow[0] ?? "").trim())
    ? actualRow.slice(1)
    : [...actualRow];
  while (normalized.length < renderedCols) {
  normalized.push("");
}
    excelRow.height = Math.max(
  24,
  ...normalized.map(c =>
    typeof c === "string"
      ? Math.min(
  Math.ceil(c.length / 45) * 15,
  120
)
      : 24
  )
);
  // fill table cells
  normalized.forEach(cell => {

  const currentCell = excelRow.getCell(col++);

  currentCell.value =
  typeof cell === "string"
    ? cell.replace(/\s*\n\s*/g, "\n")
    : cell;

  currentCell.alignment = {
    vertical: "top",
    horizontal:
  typeof cell === "string" &&
  (
    cell.length > 25 ||
    cell.split(/\s+/).length > 3
  )
    ? "left"
    : "center",
    wrapText: true,
  };
});

  // compliance + remarks
  excelRow.getCell(col++).value = "";
  excelRow.getCell(col++).value = "";

  // borders
  for (let c = 1; c < col; c++) {
    excelRow.getCell(c).border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
  }
});

  // --- Dynamic column width ---
  ws.columns = [
  { width: 8 },
  ...Array.from({ length: renderedCols }, () => ({ width: 25 })),
  { width: 22 },
  { width: 36 },
];

  ws.views = [{
  state: "frozen",
  ySplit: firstColLooksNumeric ? 2 : 1
}];
}

function populateIndexSheet(
  ws: ExcelJS.Worksheet,
  sections: Section[],
  sheetMap: Map<string, { content: string; tables: string[] }>,
) {
  ws.columns = [
    { header: "S.No.", width: 8 },
    { header: "Section #", width: 12 },
    { header: "Title", width: 50 },
    { header: "Content Sheet", width: 30 },
    { header: "Table Sheets", width: 50 },
    { header: "Source Page", width: 14 },
  ];
  styleHeaderRow(ws.getRow(1));
  sections.forEach((s, i) => {
    const m = sheetMap.get(s.number + "|" + s.title);
    ws.addRow([
      i + 1,
      s.number,
      s.title,
      m?.content ?? "",
      (m?.tables ?? []).join(", "),
      s.page,
    ]);
  });
}
function regroupTables(blocks: Block[]): Block[][] {

  const groups: Block[][] = [];

  let current: Block[] = [];

  let dominantCols = 0;

  for (const b of blocks) {

    if (b.type !== "table-row") {

      if (current.length > 0) {
        groups.push(current);
        current = [];
        dominantCols = 0;
      }

      continue;
    }

    const cols = b.cells.length;

    if (current.length === 0) {

      current.push(b);
      dominantCols = cols;
      continue;
    }

    const compatible =

  // exact
  cols === dominantCols ||

  // near match
  Math.abs(cols - dominantCols) <= 1 ||

  // continuation row
  (
    cols <= 2 &&
    current.length > 0 &&
    current[current.length - 1].cells.length >= 3 &&
    b.cells.join(" ").length < 120
  ) ||

  // long description continuation
  (
    cols >= 2 &&
    dominantCols >= 3 &&
    b.cells.some(c => c.length > 50)
  );

    if (compatible) {

      current.push(b);

      // stabilize dominant structure
      const counts =
        current.map(r => r.cells.length);

      dominantCols =
        counts
          .sort(
            (a,b) =>
              counts.filter(x => x===b).length -
              counts.filter(x => x===a).length
          )[0];

    } else {

      groups.push(current);

      current = [b];

      dominantCols = cols;
    }
  }

  if (current.length > 0) {
    groups.push(current);
  }

  return groups.filter(g => {

  if (g.length >= 2)
    return true;

  const row =
    g[0] as any;

  return (
    row?.cells?.length >= 3
  );
});
}
async function buildWorkbook(sections: Section[], images: PdfImage[]): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "RFP Extractor";
  wb.created = new Date();

  const used = new Set<string>();
  used.add("index");
  const sheetMap = new Map<string, { content: string; tables: string[] }>();

  const indexWs = wb.addWorksheet("Index", { views: [{ state: "frozen", ySplit: 1 }] });
  for (const section of sections) {
    const baseName = `${section.number}. ${section.title}`.trim();
    const contentName = safeSheetName(baseName, used);
    const ws = wb.addWorksheet(contentName, { views: [{ state: "frozen", ySplit: 1 }] });
    const regroupedTables =
  regroupTables(section.blocks);

const tableNames: string[] = [];

regroupedTables.forEach((rows, idx) => {

  const tableName = safeSheetName(
    `${section.number}. ${section.title} (${idx + 1})`,
    used
  );

  const tws =
    wb.addWorksheet(tableName, {
      views: [{ state: "frozen", ySplit: 2 }],
    });

  buildTableSheet(tws, rows);

  rows.forEach((r: any) => {
    r.__sheetName = tableName;
  });

  tableNames.push(tableName);
});

buildContentSheet(
  wb,
  ws,
  section.blocks,
  images
);

    sheetMap.set(section.number + "|" + section.title, {
      content: contentName,
      tables: tableNames,
    });
  }

  populateIndexSheet(indexWs, sections, sheetMap);

  const buf = await wb.xlsx.writeBuffer();
  return new Uint8Array(buf);
}

// --- Page range filtering ----------------------------------------------------

type PageRange = { from: number; to: number };

function parsePageRanges(input: string | null | undefined): PageRange[] {
  if (!input) return [];
  const out: PageRange[] = [];
  for (const part of input.split(/[,;]+/)) {
    const trimmed = part.trim();
    let m = trimmed.match(/^(\d+)\s*(?:-|–|—|to)\s*(\d+)$/i);
    if (!m) m = trimmed.match(/^(\d+)$/);
    if (!m) continue;
    const from = parseInt(m[1], 10);
    const to = m[2] ? parseInt(m[2], 10) : from;
    if (from > 0 && to >= from) out.push({ from, to });
  }
  return out;
}

function pageInRanges(page: number, ranges: PageRange[]): boolean {
  if (ranges.length === 0) return true;
  return ranges.some((r) => page >= r.from && page <= r.to);
}

function base64Encode(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function filterSectionsByPages(sections: Section[], ranges: PageRange[]): Section[] {
  if (ranges.length === 0) return sections;
  const out: Section[] = [];
  for (const s of sections) {
    const blocks = s.blocks.filter((b) => {
      const p = "page" in b ? b.page : 0;
      return pageInRanges(p, ranges);
    });
    if (blocks.length === 0) continue;
    out.push({ ...s, blocks });
  }
  return out;
}

// --- HTTP handler ------------------------------------------------------------

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  // TEMP: bypass auth
console.log("Auth bypassed");

  try {
    const contentType = req.headers.get("content-type") || "";
    let pdfBytes: Uint8Array;
    let filename = "rfp.pdf";
    let pageRangesInput: string | null = null;

    const url = new URL(req.url);
    pageRangesInput = url.searchParams.get("pageRanges");

    let sessionId: string | null = null;
    if (contentType.includes("multipart/form-data")) {
      const form = await req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        return new Response(JSON.stringify({ error: "No file provided" }), {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        });
      }
      pdfBytes = new Uint8Array(await file.arrayBuffer());
      if (file.name) filename = file.name;
      const pr = form.get("pageRanges");
      if (typeof pr === "string" && pr.trim()) pageRangesInput = pr;
      const sid = form.get("session_id");
      if (typeof sid === "string" && sid.trim()) {
        sessionId = sid;
      }
    } else {
      pdfBytes = new Uint8Array(await req.arrayBuffer());
    }

    if (!sessionId) {
      sessionId = req.headers.get("x-session-id")?.trim() ?? null;
    }

    if (!sessionId) {
      console.log("No session_id received → generating new one");
      sessionId = crypto.randomUUID();
    }

    console.log(`Received session_id: ${sessionId}`);

    if (SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY) {
      const { error: insertError } = await supabaseAdmin.from("rfp_sessions").insert({
        session_id: sessionId,
        file_name: filename,
        file_size: pdfBytes.length,
        status: "processed",
        metadata: { pageRanges: pageRangesInput ?? null },
      });
      if (insertError) {
        console.error("RFP session insert failed:", insertError.message);
      }

      // Upload PDF to private storage
      const storagePath = `${sessionId}/${filename}`;
      const { error: uploadError } = await supabaseAdmin.storage
        .from("rfp-files")
        .upload(storagePath, pdfBytes, {
          contentType: "application/pdf",
          upsert: false,
        });
      if (uploadError) {
        console.error("PDF upload failed:", uploadError.message);
      } else {
        console.log(`PDF uploaded to storage: ${storagePath}`);
      }
    } else {
      console.warn("SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY missing; skipping db metadata insert and storage upload.");
    }

    if (pdfBytes.length === 0) {
      return new Response(JSON.stringify({ error: "Empty file" }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const ranges = parsePageRanges(pageRangesInput);
    console.log(`Processing PDF: ${filename}, size=${pdfBytes.length}, ranges=${JSON.stringify(ranges)}`);
    const { pages } = await extractStructured(pdfBytes);

console.log(`Extracted ${pages.length} pages`);

let filteredPages = pages;

if (ranges.length > 0) {
  filteredPages = pages.filter((p) =>
  pageInRanges(p.pageNumber, ranges)
);

  console.log(
    `After page filtering: ${filteredPages.length} pages retained`
  );
}

let sections = buildSections(filteredPages);

console.log(`Detected ${sections.length} sections`);

    const images: PdfImage[] = [];

    const xlsx = await buildWorkbook(sections, images);
    const rangeSuffix = ranges.length
      ? " - p" + ranges.map((r) => (r.from === r.to ? `${r.from}` : `${r.from}-${r.to}`)).join(",")
      : "";
    const outName = filename.replace(/\.pdf$/i, "") + rangeSuffix + " - structured.xlsx";

    let signedUrl: string | null = null;
    if (SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY) {
      const excelPath = `${sessionId}/${outName}`;
      const { error: uploadError } = await supabaseAdmin.storage
        .from("rfp-files")
        .upload(excelPath, xlsx, {
          contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          upsert: true,
        });
      if (uploadError) {
        console.error("Excel upload failed:", uploadError.message);
      } else {
        const { data: urlData, error: urlError } = await supabaseAdmin.storage
          .from("rfp-files")
          .createSignedUrl(excelPath, 3600);
        if (urlError) {
          console.error("Signed URL creation failed:", urlError.message);
        } else {
          signedUrl = urlData.signedUrl;
        }
      }
    }

    const body: Record<string, unknown> = {
      fileName: outName,
      sections: sections.length,
      pages: pages.length,
      images: images.length,
    };

    if (signedUrl) {
      body.excelUrl = signedUrl;
    } else {
      body.excelBase64 = base64Encode(xlsx);
    }

    return new Response(JSON.stringify(body), {
      status: 200,
      headers: {
        ...corsHeaders,
        "Content-Type": "application/json",
      },
    });
  } catch (err) {
    console.error("process-rfp error:", err);
    const msg = err instanceof Error ? err.message : "Unknown error";
    return new Response(JSON.stringify({
  error: msg,
  debug: err
}), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
