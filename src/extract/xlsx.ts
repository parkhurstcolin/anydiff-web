/**
 * XLSX extractor — port of the engine's extractors/xlsx.py over exceljs.
 * Cell semantics mirror _CellScanner._bind_cell: a cell is kept only if it
 * has a value or a formula (styled blanks drop); formula cells carry BOTH
 * the cached result (data_only=True view) and the "=..." text
 * (data_only=False view) — the differ needs both.
 */

import ExcelJS from "exceljs";

export type CellValue = string | number | boolean | Date | null;

export interface Cell {
  addr: string;
  value: CellValue;
  formula: string | null;
}

export interface Sheet {
  name: string;
  cells: Map<string, Cell>;
}

interface Raw {
  value: CellValue;
  formula: string | null;
}

function readValue(v: unknown): CellValue {
  if (v === null || v === undefined) return null;
  const t = typeof v;
  if (t === "string" || t === "number" || t === "boolean") return v as CellValue;
  if (v instanceof Date) return v;
  if (t === "object") {
    const o = v as Record<string, unknown>;
    if (Array.isArray(o.richText)) {
      return (o.richText as { text?: string }[]).map((r) => r.text ?? "").join("");
    }
    if ("text" in o) return String(o.text); // hyperlink display text
    if ("error" in o) return String(o.error); // "#DIV/0!", openpyxl's spelling
    // unknown shape: deterministic stringify so two different unknowns can
    // never compare silently-equal — a change shows, it does not hide
    return JSON.stringify(o);
  }
  return String(v);
}

function readCell(v: unknown, ws: ExcelJS.Worksheet): Raw {
  if (v === null || v === undefined || typeof v !== "object" || v instanceof Date) {
    return { value: readValue(v), formula: null };
  }
  const o = v as Record<string, unknown>;
  if (typeof o.formula === "string") {
    // exceljs omits the leading "="; openpyxl (and the engine) keeps it
    return { value: readValue(o.result), formula: "=" + o.formula };
  }
  if (typeof o.sharedFormula === "string") {
    // Known gap: the engine expands shared formulas via openpyxl's Translator;
    // exceljs exposes only the master's address, so this reuses the master's
    // text (shared formulas are rare in hand-edited files).
    const master = ws.getCell(o.sharedFormula).value;
    const text =
      master && typeof master === "object" && "formula" in master
        ? String((master as { formula: string }).formula)
        : o.sharedFormula;
    return { value: readValue(o.result), formula: "=" + text };
  }
  return { value: readValue(v), formula: null };
}

/** .xlsx bytes -> sheets keyed by name (last duplicate wins, dict-like). */
export async function extractXlsx(
  bytes: Buffer | ArrayBuffer | Uint8Array,
): Promise<Map<string, Sheet>> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes as ArrayBuffer);
  const sheets = new Map<string, Sheet>();
  wb.eachSheet((ws) => {
    const cells = new Map<string, Cell>();
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        const { value, formula } = readCell(cell.value, ws);
        if (value === null && formula === null) return;
        cells.set(cell.address, { addr: cell.address, value, formula });
      });
    });
    sheets.set(ws.name, { name: ws.name, cells });
  });
  return sheets;
}
