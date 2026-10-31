/**
 * Opt-in formula recalculation (browser-only, clearly labelled in the UI):
 * HyperFormula re-evaluates every formula cell, and the diff then compares
 * COMPUTED values instead of the results Excel last saved. This is the
 * honest answer to files written by tools that never calculate (pandas,
 * openpyxl re-saves), at the cost of not matching "what the file says".
 *
 * Parity note: the Python engine never recalculates, so no golden depends
 * on this file — the differ just receives different input values and the
 * UI marks the run as computed.
 */

import { HyperFormula } from "hyperformula";

import type { Cell, Sheet } from "../extract/xlsx";

const MAX_RECALC_CELLS = 200_000; // sparse-huge guard: skip instead of hanging

export interface RecalcResult {
  sheets: Map<string, Sheet>;
  computed: number; // formula cells that got a fresh value
  failed: number; // formula cells kept at their saved value (error/date/etc.)
  skipped: string | null; // set when nothing was recalculated, with the reason
}

const splitAddr = (addr: string): [number, number] | null => {
  const m = /^([A-Z]+)(\d+)$/.exec(addr);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]), col];
};

const extent = (sheet: Sheet): [number, number] => {
  let r = 0;
  let c = 0;
  for (const addr of sheet.cells.keys()) {
    const rc = splitAddr(addr);
    if (rc) {
      r = Math.max(r, rc[0]);
      c = Math.max(c, rc[1]);
    }
  }
  return [Math.max(1, r), Math.max(1, c)];
};

export function recalcSheets(sheets: Map<string, Sheet>): RecalcResult {
  const passthrough = (skipped: string): RecalcResult => ({ sheets, computed: 0, failed: 0, skipped });

  let cells = 0;
  for (const s of sheets.values()) {
    const [r, c] = extent(s);
    cells += r * c;
  }
  if (cells > MAX_RECALC_CELLS) return passthrough("workbook too large to recalculate");

  // engine sheet contents: raw values, formula strings kept with their "="
  const build: Record<string, unknown[][]> = {};
  for (const [name, s] of sheets) {
    const [rows, cols] = extent(s);
    const grid: unknown[][] = Array.from({ length: rows }, () => Array<unknown>(cols).fill(null));
    for (const cell of s.cells.values()) {
      const rc = splitAddr(cell.addr);
      if (!rc) continue;
      if (cell.formula !== null) grid[rc[0] - 1][rc[1] - 1] = cell.formula;
      else if (cell.value instanceof Date) grid[rc[0] - 1][rc[1] - 1] = cell.value.getTime(); // engine has no date type
      else grid[rc[0] - 1][rc[1] - 1] = cell.value;
    }
    build[name] = grid;
  }

  let hf: HyperFormula;
  try {
    hf = HyperFormula.buildFromSheets(build as Parameters<typeof HyperFormula.buildFromSheets>[0], { licenseKey: "gpl-v3" });
  } catch {
    return passthrough("engine could not parse the workbook");
  }

  let computed = 0;
  let failed = 0;
  const out = new Map<string, Sheet>();
  try {
    let sheetId = -1;
    for (const [name, s] of sheets) {
      sheetId += 1; // build order == engine sheet ids
      const id = sheetId;
      const cells2 = new Map<string, Cell>();
      for (const [addr, cell] of s.cells) {
        if (cell.formula === null) {
          cells2.set(addr, cell);
          continue;
        }
        const rc = splitAddr(addr);
        if (!rc) {
          failed += 1;
          cells2.set(addr, cell);
          continue;
        }
        const raw = hf.getCellValue({ sheet: id, row: rc[0] - 1, col: rc[1] - 1 });
        const isError = raw instanceof Error || (typeof raw === "object" && raw !== null && "value" in raw && String((raw as { value: unknown }).value).startsWith("#"));
        if (isError) {
          failed += 1; // the engine could not evaluate it — keep the saved value
          cells2.set(addr, cell);
        } else if (raw === null || typeof raw === "number" || typeof raw === "string" || typeof raw === "boolean") {
          // a saved Date vs a computed serial is not a comparable pair: keep saved
          if (cell.value instanceof Date && typeof raw === "number") {
            failed += 1;
            cells2.set(addr, cell);
          } else {
            computed += 1;
            cells2.set(addr, { ...cell, value: raw });
          }
        } else {
          failed += 1; // arrays and exotic results stay as saved
          cells2.set(addr, cell);
        }
      }
      out.set(name, { ...s, cells: cells2 });
    }
  } finally {
    hf.destroy();
  }
  return { sheets: out, computed, failed, skipped: null };
}
