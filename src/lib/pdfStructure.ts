/**
 * pdfStructure.ts
 *
 * Browser-side port of the edge function's structured PDF extraction.
 * Produces the same multi-sheet Excel as the server:
 *   Index sheet + one content sheet per section + one table sheet per table group.
 *
 * Uses: pdfjs-dist (already in project) + xlsx (SheetJS, already in project)
 */

import * as XLSX from "xlsx";
import pdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type ParagraphBlock = { type: "paragraph"; text: string; page: number };
type SubheadingBlock = { type: "subheading"; text: string; level: number; page: number };
type TableRowBlock = { type: "table-row"; text: string; cells: string[]; page: number; tableId: number };

type Block = ParagraphBlock | SubheadingBlock | TableRowBlock;

type Section = {
  number: string;
  title: string;
  page: number;
  blocks: Block[];
};

type PageRange = { from: number; to: number };

// ---------------------------------------------------------------------------
// PDF text extraction (pdfjs-dist)
// ---------------------------------------------------------------------------

export async function extractPages(file: File): Promise<{ pages: string[][]; pageCount: number }> {
  const pdfjsLib = await import("pdfjs-dist/legacy/build/pdf");
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pageCount = pdf.numPages;
  const pages: string[][] = [];

  for (let p = 1; p <= pageCount; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();

    // Join items into a raw string then split into lines, same as the edge fn
    const raw = content.items.map((item: any) => item.str).join("\n");
    const lines = raw.split(/\r?\n/).map((l: string) => l.trim());
    const out: string[] = [];
    let prevBlank = true;
    for (const l of lines) {
      if (l.length === 0) {
        if (!prevBlank) out.push("");
        prevBlank = true;
      } else {
        out.push(l);
        prevBlank = false;
      }
    }
    pages.push(out);
  }

  return { pages, pageCount };
}

// ---------------------------------------------------------------------------
// Structure detection (direct port from edge function)
// ---------------------------------------------------------------------------

const TOP_HEADING_RE =
  /^(?:section|chapter|part)\s+(\d{1,2})[.\):–\-—\s]+(.{2,80}?)$|^(\d{1,2})[.)]\s+([A-Z][A-Za-z0-9 &/,\-—:]{2,80})$/i;
const SUB_HEADING_RE = /^(\d{1,2}(?:\.\d{1,2}){1,3})\s+([A-Z].{1,100}?)$/;

function splitMultiGap(line: string): string[] | null {
  if (line.includes("\t")) {
    const parts = line.split(/\t+/).map((s) => s.trim()).filter(Boolean);
    return parts.length >= 2 ? parts : null;
  }
  const parts = line.split(/\s{2,}/).map((s) => s.trim()).filter(Boolean);
  if (parts.length < 3 || parts.length > 8) return null;
  if (line.length > 160) return null;
  if (/[.!?]$/.test(line)) return null;
  const anyLong = parts.some((p) => p.split(/\s+/).length > 8 || p.length > 50);
  if (anyLong) return null;
  return parts;
}

function splitNumberedRow(line: string): string[] | null {
  const m = line.match(/^(\d{1,3})[.)]?\s+(.+)$/);
  if (!m) return null;
  const rest = m[2].trim();
  if (rest.length > 120) return null;
  if (/[.!?]$/.test(rest)) return null;
  const multi = rest.split(/\s{2,}/).map((s) => s.trim()).filter(Boolean);
  if (multi.length >= 2) return [m[1], ...multi];
  return [m[1], rest];
}

function mergeParagraphs(blocks: Block[]): Block[] {
  const BREAK = "\u0000BREAK";
  const out: Block[] = [];
  let buf: ParagraphBlock | null = null;

  const flush = () => { if (buf) { out.push(buf); buf = null; } };

  for (const b of blocks) {
    if (b.type === "paragraph") {
      if (b.text === BREAK) { flush(); continue; }
      if (buf) {
        const prev = buf.text.trimEnd();
        const endsSentence = /[.!?:](?:[\"'\)\]]?)$/.test(prev);
        const startsExplicit = /^(?:[-•·*]|\(?\\d+[.)]\s+)/.test(b.text);
        if (endsSentence || startsExplicit) {
          flush();
          buf = { ...b };
        } else {
          const joined = prev.endsWith("-")
            ? prev.slice(0, -1) + b.text
            : prev + " " + b.text;
          buf = { ...buf, text: joined };
        }
      } else {
        buf = { ...b };
      }
    } else {
      flush();
      out.push(b);
    }
  }
  flush();
  return out;
}

function buildSections(pages: string[][]): Section[] {
  const BREAK = "\u0000BREAK";
  const sections: Section[] = [];
  let current: Section | null = null;
  let nextTableId = 1;
  let activeTableId: number | null = null;
  let tableRunCount = 0;

  const pushIntro = (page: number) => {
    if (!current) {
      current = { number: "0", title: "Preamble", page, blocks: [] };
      sections.push(current);
    }
  };

  const closeTableRun = () => { activeTableId = null; tableRunCount = 0; };

  pages.forEach((lines, pageIdx) => {
    const page = pageIdx + 1;
    for (const raw of lines) {
      if (raw.length === 0) {
        if (current?.blocks.length) {
          const last = current.blocks[current.blocks.length - 1];
          if (last.type === "paragraph" && last.text !== BREAK) {
            current.blocks.push({ type: "paragraph", text: BREAK, page });
          }
        }
        if (tableRunCount >= 2) closeTableRun();
        continue;
      }
      const line = raw.replace(/\s+/g, " ").trim();
      if (!line) continue;

      const top = line.match(TOP_HEADING_RE);
      if (top) {
        closeTableRun();
        current = {
          number: ((top[1] ?? top[3]) as string),
          title: ((top[2] ?? top[4]) as string).trim(),
          page,
          blocks: [],
        };
        sections.push(current);
        continue;
      }

      const sub = line.match(SUB_HEADING_RE);
      if (sub) {
        closeTableRun();
        pushIntro(page);
        current!.blocks.push({
          type: "subheading",
          text: `${sub[1]} ${sub[2].trim()}`,
          level: sub[1].split(".").length,
          page,
        });
        continue;
      }

      const gapCells = splitMultiGap(line);
      if (gapCells) {
        pushIntro(page);
        if (activeTableId === null) { activeTableId = nextTableId++; tableRunCount = 0; }
        tableRunCount++;
        current!.blocks.push({ type: "table-row", text: line, cells: gapCells, page, tableId: activeTableId });
        continue;
      }

      const numCells = splitNumberedRow(line);
      if (numCells) {
        pushIntro(page);
        if (activeTableId === null) { activeTableId = nextTableId++; tableRunCount = 0; }
        tableRunCount++;
        current!.blocks.push({ type: "table-row", text: line, cells: numCells, page, tableId: activeTableId });
        continue;
      }

      // Demote short table runs back to paragraphs
      if (activeTableId !== null && tableRunCount < 2) {
        for (let i = current!.blocks.length - 1; i >= 0; i--) {
          const b = current!.blocks[i];
          if (b.type === "table-row" && b.tableId === activeTableId) {
            current!.blocks[i] = { type: "paragraph", text: b.text, page: b.page };
          } else break;
        }
      }
      closeTableRun();

      pushIntro(page);
      current!.blocks.push({ type: "paragraph", text: line, page });
    }
    if (tableRunCount < 2) closeTableRun();
  });

  // Final cleanup: demote orphan table groups (<2 rows) and merge paragraphs
  for (const s of sections) {
    const tableCounts = new Map<number, number>();
    for (const b of s.blocks) {
      if (b.type === "table-row") tableCounts.set(b.tableId, (tableCounts.get(b.tableId) ?? 0) + 1);
    }
    s.blocks = s.blocks.map((b) => {
      if (b.type === "table-row" && (tableCounts.get(b.tableId) ?? 0) < 2) {
        return { type: "paragraph", text: b.text, page: b.page } as ParagraphBlock;
      }
      return b;
    });
    s.blocks = mergeParagraphs(s.blocks).filter(
      (b) => !(b.type === "paragraph" && b.text === BREAK)
    );
  }

  return sections;
}

// ---------------------------------------------------------------------------
// Page-range filtering
// ---------------------------------------------------------------------------

export function parsePageRanges(input: string | null | undefined): PageRange[] {
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

function filterSectionsByPages(sections: Section[], ranges: PageRange[]): Section[] {
  if (ranges.length === 0) return sections;
  return sections
    .map((s) => ({
      ...s,
      blocks: s.blocks.filter((b) => {
        const p = "page" in b ? b.page : 0;
        return ranges.some((r) => p >= r.from && p <= r.to);
      }),
    }))
    .filter((s) => s.blocks.length > 0);
}

// ---------------------------------------------------------------------------
// Excel builder (SheetJS / xlsx)
// ---------------------------------------------------------------------------

function safeSheetName(name: string, used: Set<string>): string {
  let n = name.replace(/[:\\/?*[\]]/g, " ").replace(/\s+/g, " ").trim();
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

function buildContentSheetData(
  blocks: Block[],
  tableNameByTableId: Map<number, string>
): (string | number)[][] {
  const rows: (string | number)[][] = [
    ["S.No.", "Technical Specifications", "Compliance (Yes/No)", "Remarks"],
  ];
  const seenTableMarkers = new Set<number>();
  let sno = 1;

  for (const b of blocks) {
    if (b.type === "table-row") {
      if (!seenTableMarkers.has(b.tableId)) {
        seenTableMarkers.add(b.tableId);
        const sheetName = tableNameByTableId.get(b.tableId);
        const text = sheetName
          ? `→ See table on sheet: "${sheetName}"`
          : "→ Table content (see adjacent table sheet)";
        rows.push([sno++, text, "", ""]);
      }
      continue;
    }
    if (b.type === "subheading") {
      rows.push([sno++, b.text, "", ""]);
      continue;
    }
    // paragraph
    rows.push([sno++, b.text, "", ""]);
  }

  if (sno === 1) {
    rows.push([1, "(No prose content detected for this section.)", "", ""]);
  }
  return rows;
}

function buildTableSheetData(
  tableRows: Extract<Block, { type: "table-row" }>[]
): (string | number)[][] {
  const cellRows = tableRows.map((r) => r.cells);
  const pageByRow = tableRows.map((r) => r.page);
  const maxCols = Math.max(2, ...cellRows.map((r) => r.length));

  let headerRow: string[] | null = null;
  if (cellRows.length > 0) {
    const first = cellRows[0];
    if (first.every((c) => c.length <= 40 && !/\.$/.test(c))) {
      headerRow = first;
    }
  }

  const dataRows = headerRow ? cellRows.slice(1) : cellRows;
  const dataPages = headerRow ? pageByRow.slice(1) : pageByRow;
  const headers = [
    "S.No.",
    ...(headerRow ?? Array.from({ length: maxCols }, (_, i) => `Col ${i + 1}`)),
  ];
  while (headers.length < maxCols + 1) headers.push(`Col ${headers.length}`);

  const result: (string | number)[][] = [
    [...headers.slice(0, maxCols + 1), "Page", "Compliance (Yes/No)", "Remarks"],
  ];

  dataRows.forEach((row, i) => {
    const padded = [...row];
    while (padded.length < maxCols) padded.push("");
    result.push([i + 1, ...padded, dataPages[i] ?? "", "", ""]);
  });

  return result;
}

function applyColWidths(ws: XLSX.WorkSheet, widths: number[]) {
  ws["!cols"] = widths.map((wch) => ({ wch }));
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export type ExtractionResult = {
  url: string;
  name: string;
  sections: number;
  pages: number;
};

export async function extractStructuredExcel(
  file: File,
  pageRangesInput: string
): Promise<ExtractionResult> {
  const { pages, pageCount } = await extractPages(file);
  const ranges = parsePageRanges(pageRangesInput);

  let sections = buildSections(pages);
  if (ranges.length > 0) sections = filterSectionsByPages(sections, ranges);

  const wb = XLSX.utils.book_new();
  const used = new Set<string>();
  used.add("index");

  // Index sheet placeholder data — we'll fill it after all sections are named
  const indexRows: (string | number)[][] = [
    ["S.No.", "Section #", "Title", "Content Sheet", "Table Sheets", "Source Page"],
  ];

  for (const [si, section] of sections.entries()) {
    const baseName = `${section.number}. ${section.title}`.trim();
    const contentName = safeSheetName(baseName, used);

    // Collect table groups
    const tableGroups = new Map<number, Extract<Block, { type: "table-row" }>[]>();
    for (const b of section.blocks) {
      if (b.type === "table-row") {
        if (!tableGroups.has(b.tableId)) tableGroups.set(b.tableId, []);
        tableGroups.get(b.tableId)!.push(b);
      }
    }

    // Build table sheets
    const tableNameByTableId = new Map<number, string>();
    const tableSheetNames: string[] = [];
    let tIdx = 1;
    for (const [tableId, rows] of tableGroups) {
      if (rows.length < 2) continue;
      const suffix = tableGroups.size > 1 ? ` — Table ${tIdx}` : " — Table";
      const tableName = safeSheetName(`${section.number}. ${section.title}${suffix}`, used);
      const tableData = buildTableSheetData(rows);
      const tws = XLSX.utils.aoa_to_sheet(tableData);
      applyColWidths(tws, [8, ...Array(tableData[0].length - 4).fill(24), 10, 22, 36]);
      XLSX.utils.book_append_sheet(wb, tws, tableName);
      tableNameByTableId.set(tableId, tableName);
      tableSheetNames.push(tableName);
      tIdx++;
    }

    // Build content sheet
    const contentData = buildContentSheetData(section.blocks, tableNameByTableId);
    const cws = XLSX.utils.aoa_to_sheet(contentData);
    applyColWidths(cws, [8, 70, 22, 36]);
    XLSX.utils.book_append_sheet(wb, cws, contentName);

    // Index row
    indexRows.push([
      si + 1,
      section.number,
      section.title,
      contentName,
      tableSheetNames.join(", "),
      section.page,
    ]);
  }

  // Add Index sheet first by rebuilding — SheetJS doesn't reorder easily,
  // so we prepend by creating a new workbook with Index first.
  const finalWb = XLSX.utils.book_new();
  const indexWs = XLSX.utils.aoa_to_sheet(indexRows);
  applyColWidths(indexWs, [8, 12, 50, 30, 50, 14]);
  XLSX.utils.book_append_sheet(finalWb, indexWs, "Index");
  for (const name of wb.SheetNames) {
    XLSX.utils.book_append_sheet(finalWb, wb.Sheets[name], name);
  }

  const rangeSuffix =
    ranges.length
      ? " - p" + ranges.map((r) => (r.from === r.to ? `${r.from}` : `${r.from}-${r.to}`)).join(",")
      : "";
  const outName = file.name.replace(/\.pdf$/i, "") + rangeSuffix + " - structured.xlsx";

  const buf = XLSX.write(finalWb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return {
    url: URL.createObjectURL(blob),
    name: outName,
    sections: sections.length,
    pages: pageCount,
  };
}
