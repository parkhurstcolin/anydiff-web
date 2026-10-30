/**
 * Side-by-side view model (display-only — never feeds the diff, never
 * touches the Change contract). Line-up rules MIRROR the differ's own
 * alignment (same blockTokens, same sequenceMatcher) so a highlight here
 * can never contradict a row in the change list.
 */

import { blockToken, fmtSig } from "./docxDiff";
import type { Change, Confidence } from "./model";
import { sequenceMatcher } from "./sequenceMatcher";
import { paragraphText, type Block, type Docx, type Run, type TableBlock } from "../extract/docx";
import type { Cell, Sheet } from "../extract/xlsx";
import { fmtValue } from "./format";

/** What the extractors produced, kept around so the UI can render the
 *  documents themselves, not just the change list. */
export type DocPair =
  | { kind: "xlsx"; old: Map<string, Sheet>; new: Map<string, Sheet> }
  | { kind: "docx"; old: Docx; new: Docx };

// ---------------------------------------------------------------- XLSX ----

export interface GridCell {
  text: string;
  title: string | null;
  changed: boolean;
}

export interface GridRow {
  n: number;
  cells: (GridCell | null)[]; // index = column offset from col0
}

export interface SheetPane {
  kind: "pair" | "removed" | "added";
  oldLabel: string; // "" when no old side
  newLabel: string; // "" when no new side
  rename: { old: string; new: string } | null;
  confidence: Confidence;
  row0: number;
  col0: number;
  oldRows: GridRow[];
  newRows: GridRow[];
  truncated: boolean;
  note: string | null; // e.g. "sheet removed", "sheet added"
}

const MAX_GRID_ROWS = 80;
const MAX_GRID_COLS = 26;
const CONTEXT = 2;

const colLetter = (n: number): string => {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

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

function gridCell(c: Cell, changed: boolean): GridCell {
  if (c.formula !== null) {
    return {
      text: c.formula,
      title: c.value === null ? null : `shows ${fmtValue(c.value)}`,
      changed,
    };
  }
  return { text: fmtValue(c.value), title: null, changed };
}

function gridFor(sheet: Sheet | undefined, changedAddrs: Set<string>, row0: number, row1: number, col0: number, col1: number): GridRow[] {
  const rows: GridRow[] = [];
  for (let r = row0; r <= row1; r++) {
    const cells: (GridCell | null)[] = [];
    for (let c = col0; c <= col1; c++) {
      const addr = colLetter(c) + r;
      const cell = sheet?.cells.get(addr);
      cells.push(cell ? gridCell(cell, changedAddrs.has(addr)) : null);
    }
    rows.push({ n: r, cells });
  }
  return rows;
}

/** Panes for the spreadsheet side-by-side: one per new sheet (workbook
 *  order), removed sheets last. Windows center on the sheet's changed
 *  addresses with ±CONTEXT rows/cols of breathing room. */
export function sideXlsx(changes: Change[], oldSheets: Map<string, Sheet>, newSheets: Map<string, Sheet>): SheetPane[] {
  const renamed = new Map<string, { old: string; conf: Confidence }>();
  const removed = new Set<string>();
  const added = new Set<string>();
  const cellByAddr = new Map<string, Change>(); // `${sheet}!${addr}`
  for (const c of changes) {
    if (c.kind === "sheet_renamed") renamed.set(c.where, { old: String(c.before), conf: c.confidence });
    else if (c.kind === "sheet_removed") removed.add(c.where);
    else if (c.kind === "sheet_added") added.add(c.where);
    else if (c.kind === "cell" && c.where.includes("!")) cellByAddr.set(c.where, c);
  }

  const panes: SheetPane[] = [];
  const build = (kind: SheetPane["kind"], oldName: string, newName: string, rename: SheetPane["rename"], confidence: Confidence, note: string | null) => {
    const o = oldSheets.get(oldName);
    const n = newSheets.get(newName);
    const changed = new Set<string>();
    for (const key of cellByAddr.keys()) {
      if (key.startsWith(newName + "!")) changed.add(key.slice(newName.length + 1));
    }
    let r0 = 1;
    let c0 = 1;
    let r1 = 12;
    let c1 = 10;
    if (changed.size) {
      let minR = 1e9;
      let maxR = 0;
      let minC = 1e9;
      let maxC = 0;
      for (const addr of changed) {
        const rc = splitAddr(addr);
        if (!rc) continue;
        minR = Math.min(minR, rc[0]);
        maxR = Math.max(maxR, rc[0]);
        minC = Math.min(minC, rc[1]);
        maxC = Math.max(maxC, rc[1]);
      }
      r0 = Math.max(1, minR - CONTEXT);
      r1 = maxR + CONTEXT;
      c0 = Math.max(1, minC - CONTEXT);
      c1 = maxC + CONTEXT;
    } else {
      const [er, ec] = extent(o ?? n!);
      r1 = Math.min(er, r1);
      c1 = Math.min(ec, c1);
    }
    const truncated = r1 - r0 + 1 > MAX_GRID_ROWS || c1 - c0 + 1 > MAX_GRID_COLS;
    r1 = Math.min(r1, r0 + MAX_GRID_ROWS - 1);
    c1 = Math.min(c1, c0 + MAX_GRID_COLS - 1);
    panes.push({
      kind,
      oldLabel: kind === "added" ? "" : oldName,
      newLabel: kind === "removed" ? "" : newName,
      rename,
      confidence,
      row0: r0,
      col0: c0,
      oldRows: kind === "added" ? [] : gridFor(o, changed, r0, r1, c0, c1),
      newRows: kind === "removed" ? [] : gridFor(n, changed, r0, r1, c0, c1),
      truncated,
      note,
    });
  };

  for (const name of newSheets.keys()) {
    if (added.has(name)) {
      build("added", "", name, null, "exact", "sheet added");
      continue;
    }
    const r = renamed.get(name);
    build("pair", r ? r.old : name, name, r ? { old: r.old, new: name } : null, r?.conf ?? "exact", null);
  }
  for (const name of removed) build("removed", name, "", null, "exact", "sheet removed");
  return panes;
}

// ---------------------------------------------------------------- DOCX ----

export interface TableRowPair {
  status: "same" | "edit" | "removed" | "added";
  oldCells: string[] | null;
  newCells: string[] | null;
}

export type DocLine =
  | { type: "same"; runs: Run[]; style: string }
  | { type: "removed"; runs: Run[]; style: string }
  | { type: "added"; runs: Run[]; style: string }
  | { type: "edit"; oldText: string; newText: string; oldStyle: string; newStyle: string; styleChanged: boolean; fmtChanged: boolean }
  | { type: "tablePair"; rows: TableRowPair[] }
  | { type: "tableDump"; side: "old" | "new"; table: TableBlock };

const rowKey = (row: string[]) => row.join("\x1e");

function tableRows(oldt: TableBlock, newt: TableBlock): TableRowPair[] {
  const out: TableRowPair[] = [];
  for (const [tag, i1, i2, j1, j2] of sequenceMatcher(oldt.rows.map(rowKey), newt.rows.map(rowKey))) {
    if (tag === "equal") {
      for (let k = 0; k < i2 - i1; k++) out.push({ status: "same", oldCells: oldt.rows[i1 + k], newCells: newt.rows[j1 + k] });
    } else if (tag === "replace" && i2 - i1 === j2 - j1) {
      for (let k = 0; k < i2 - i1; k++) out.push({ status: "edit", oldCells: oldt.rows[i1 + k], newCells: newt.rows[j1 + k] });
    } else {
      for (let k = i1; k < i2; k++) out.push({ status: "removed", oldCells: oldt.rows[k], newCells: null });
      for (let k = j1; k < j2; k++) out.push({ status: "added", oldCells: null, newCells: newt.rows[k] });
    }
  }
  return out;
}

const isPara = (b: Block): b is Extract<Block, { blockType: "paragraph" }> => b.blockType === "paragraph";

function pairLines(oldb: Block, newb: Block): DocLine[] {
  if (oldb.blockType === "table" && newb.blockType === "table") {
    return [{ type: "tablePair", rows: tableRows(oldb, newb) }];
  }
  if (isPara(oldb) && isPara(newb)) {
    const oldText = paragraphText(oldb);
    const newText = paragraphText(newb);
    return [
      {
        type: "edit",
        oldText,
        newText,
        oldStyle: oldb.style,
        newStyle: newb.style,
        styleChanged: oldb.style !== newb.style,
        fmtChanged: oldb.style === newb.style && fmtSig(oldb) !== fmtSig(newb),
      },
    ];
  }
  const out: DocLine[] = [];
  if (oldb.blockType === "table") out.push({ type: "tableDump", side: "old", table: oldb });
  else out.push({ type: "removed", runs: oldb.runs, style: oldb.style });
  if (newb.blockType === "table") out.push({ type: "tableDump", side: "new", table: newb });
  else out.push({ type: "added", runs: newb.runs, style: newb.style });
  return out;
}

const paraLines = (blocks: Block[], i1: number, i2: number, type: "removed" | "added"): DocLine[] =>
  Array.from({ length: i2 - i1 }, (_, k) => {
    const b = blocks[i1 + k];
    return b.blockType === "table"
      ? ({ type: "tableDump", side: type === "removed" ? "old" : "new", table: b } as DocLine)
      : ({ type, runs: b.runs, style: b.style } as DocLine);
  });

/** One entry per line-row; equal blocks collapse into a single `same` line
 *  (rendered on both sides) — the git "context" rows. */
export function sideDocx(old: Docx, new_: Docx): DocLine[] {
  const lines: DocLine[] = [];
  for (const [tag, i1, i2, j1, j2] of sequenceMatcher(old.blocks.map(blockToken), new_.blocks.map(blockToken))) {
    if (tag === "equal") {
      for (let k = 0; k < i2 - i1; k++) {
        const b = old.blocks[i1 + k];
        if (b.blockType === "table") {
          // equal block tokens guarantee the new side is a table too
          lines.push({ type: "tablePair", rows: tableRows(b, new_.blocks[j1 + k] as TableBlock) });
        } else {
          lines.push({ type: "same", runs: b.runs, style: b.style });
        }
      }
    } else if (tag === "replace" && i2 - i1 === j2 - j1) {
      for (let k = 0; k < i2 - i1; k++) lines.push(...pairLines(old.blocks[i1 + k], new_.blocks[j1 + k]));
    } else if (tag === "replace" && i2 - i1 === 1 && j2 - j1 > 1) {
      // same split-paragraph reading the differ gives (edit + additions)
      lines.push(...pairLines(old.blocks[i1], new_.blocks[j1]));
      lines.push(...paraLines(new_.blocks, j1 + 1, j2, "added"));
    } else {
      if (i2 > i1) lines.push(...paraLines(old.blocks, i1, i2, "removed"));
      if (j2 > j1) lines.push(...paraLines(new_.blocks, j1, j2, "added"));
    }
  }
  return lines;
}

// ------------------------------------------------------------ targets ----

/** Where the navigator jumps: a highlighted cell (pane-local grid indices,
 *  identical on both halves) or, at row/col -1, the pane header (a sheet
 *  that was renamed / added / removed). */
export interface XlsxTarget {
  pane: number;
  row: number;
  col: number;
}

export function xlsxTargets(panes: SheetPane[]): XlsxTarget[] {
  const out: XlsxTarget[] = [];
  panes.forEach((p, pi) => {
    if (p.rename || p.note) out.push({ pane: pi, row: -1, col: -1 });
    const rows = p.newRows.length ? p.newRows : p.oldRows; // removed panes only have the old grid
    rows.forEach((r, ri) => r.cells.forEach((c, ci) => c?.changed && out.push({ pane: pi, row: ri, col: ci })));
  });
  return out;
}

/** Line indices worth jumping to: anything that is not plain context
 *  (tablePair blocks whose rows are all equal count as context). */
export function docxTargets(lines: DocLine[]): number[] {
  const out: number[] = [];
  lines.forEach((l, i) => {
    if (l.type === "same") return;
    if (l.type === "tablePair" && l.rows.every((r) => r.status === "same")) return;
    out.push(i);
  });
  return out;
}
