/**
 * DOCX differ — port of diff.py's docx section: block-token SequenceMatcher
 * opcodes -> paragraph edits / table row pairing / honest add-remove dumps.
 * The split-paragraph, size-changing-table and refusal behaviors are all
 * golden-tested against the Python engine.
 */

import { jaccard } from "./align";
import { change, type Change } from "./model";
import { sequenceMatcher } from "./sequenceMatcher";
import {
  paragraphText,
  type Block,
  type Docx,
  type ParagraphBlock,
  type TableBlock,
} from "../extract/docx";

const blockToken = (b: Block): string =>
  b.blockType === "table"
    ? "T|" + b.rows.map((row) => row.join("\x1e")).join("\x1f")
    : `P|${b.style}|` +
      b.runs.map((r) => `${r.text}\x1d${r.bold ? 1 : 0}${r.italic ? 1 : 0}`).join("\x1e");

const rowKey = (row: string[]) => row.join("\x1e");
const joinRow = (row: string[]) => row.join(" | ");

/** Pair rows inside a size-changing replace block: equal first cells anchor
 *  (row labels), the rest pairs greedily by Jaccard >= 0.5 of cell sets;
 *  unpaired rows honestly stay removed/added (mirrors align.py anchoring). */
function pairRows(oldt: TableBlock, newt: TableBlock, i1: number, i2: number, j1: number, j2: number) {
  const pairs = new Map<number, number>();
  const usedOld = new Set<number>();
  const usedNew = new Set<number>();
  for (let oi = i1; oi < i2; oi++) {
    if (!oldt.rows[oi].length) continue;
    const label = oldt.rows[oi][0];
    for (let nj = j1; nj < j2; nj++) {
      if (usedNew.has(nj) || !newt.rows[nj].length) continue;
      if (newt.rows[nj][0] === label) {
        pairs.set(oi, nj);
        usedOld.add(oi);
        usedNew.add(nj);
        break;
      }
    }
  }
  const sims: [number, number, number][] = [];
  for (let oi = i1; oi < i2; oi++) {
    for (let nj = j1; nj < j2; nj++) {
      if (usedOld.has(oi) || usedNew.has(nj)) continue;
      sims.push([oi, nj, jaccard(new Set(oldt.rows[oi]), new Set(newt.rows[nj]))]);
    }
  }
  sims.sort((x, y) => y[2] - x[2] || x[0] - y[0] || x[1] - y[1]);
  for (const [oi, nj, sim] of sims) {
    if (sim < 0.5 || usedOld.has(oi) || usedNew.has(nj)) continue;
    pairs.set(oi, nj);
    usedOld.add(oi);
    usedNew.add(nj);
  }
  return pairs;
}

const paraWhere = (p: ParagraphBlock, n: number) =>
  p.style !== "Normal" ? `¶${n} [${p.style}]` : `¶${n}`;

const fmtSig = (p: ParagraphBlock) => p.runs.map((r) => `${r.text}|${r.bold}|${r.italic}`).join(";");

function blockAdded(display: string, b: Block, n: number): Change {
  if (b.blockType === "table") {
    return change({
      path: display,
      kind: "row",
      where: `table ${n} row 1`,
      before: null,
      after: b.rows.length ? joinRow(b.rows[0]) : "",
      note: `table added (${b.rows.length} rows)`,
    });
  }
  return change({
    path: display,
    kind: "paragraph",
    where: paraWhere(b, n),
    before: null,
    after: paragraphText(b),
    note: b.style === "Normal" ? null : `added ${b.style} paragraph`,
  });
}

function blockRemoved(display: string, b: Block, n: number): Change {
  if (b.blockType === "table") {
    return change({
      path: display,
      kind: "row",
      where: `table ${n} row 1`,
      before: b.rows.length ? joinRow(b.rows[0]) : "",
      after: null,
      note: `table removed (${b.rows.length} rows)`,
    });
  }
  return change({
    path: display,
    kind: "paragraph",
    where: paraWhere(b, n),
    before: paragraphText(b),
    after: null,
    note: b.style === "Normal" ? null : `removed ${b.style} paragraph`,
  });
}

function paraChange(display: string, oldp: ParagraphBlock, newp: ParagraphBlock, n: number): Change | null {
  let note: string | null = null;
  if (paragraphText(oldp) !== paragraphText(newp)) {
    if (oldp.style !== newp.style) note = `style: ${oldp.style} -> ${newp.style}`;
  } else if (oldp.style !== newp.style) {
    note = `style: ${oldp.style} -> ${newp.style}`;
  } else if (fmtSig(oldp) !== fmtSig(newp)) {
    note = "run formatting changed";
  } else {
    return null;
  }
  return change({
    path: display,
    kind: "paragraph",
    where: paraWhere(newp, n),
    before: paragraphText(oldp),
    after: paragraphText(newp),
    note,
  });
}

function tableRowChanges(
  display: string,
  label: string,
  oldt: TableBlock,
  newt: TableBlock,
  confidence: "exact" | "moved" | "manual-inspect" = "exact",
): Change[] {
  const changes: Change[] = [];
  const row = (kind: "row", where: string, before: string | null, after: string | null) =>
    changes.push(change({ path: display, kind, where, before, after, confidence }));
  for (const [tag, i1, i2, j1, j2] of sequenceMatcher(oldt.rows.map(rowKey), newt.rows.map(rowKey))) {
    if (tag === "equal") continue;
    const nOld = i2 - i1;
    const nNew = j2 - j1;
    if (tag === "replace" && nOld === nNew) {
      for (let k = 0; k < nOld; k++) {
        row("row", `${label} row ${j1 + k + 1}`, joinRow(oldt.rows[i1 + k]), joinRow(newt.rows[j1 + k]));
      }
    } else if (tag === "replace") {
      const pairs = pairRows(oldt, newt, i1, i2, j1, j2);
      const pairedNew = new Map<number, number>();
      for (const [oi, nj] of pairs) pairedNew.set(nj, oi);
      for (let nj = j1; nj < j2; nj++) {
        if (pairedNew.has(nj)) {
          row("row", `${label} row ${nj + 1}`, joinRow(oldt.rows[pairedNew.get(nj)!]), joinRow(newt.rows[nj]));
        } else {
          row("row", `${label} row ${nj + 1}`, null, joinRow(newt.rows[nj]));
        }
      }
      for (let oi = i1; oi < i2; oi++) {
        if (!pairs.has(oi)) row("row", `${label} row ${oi + 1}`, joinRow(oldt.rows[oi]), null);
      }
    } else {
      for (let k = 0; k < nOld; k++) row("row", `${label} row ${i1 + k + 1}`, joinRow(oldt.rows[i1 + k]), null);
      for (let k = 0; k < nNew; k++) row("row", `${label} row ${j1 + k + 1}`, null, joinRow(newt.rows[j1 + k]));
    }
  }
  return changes;
}

const emptyTable = (): TableBlock => ({ blockType: "table", rows: [] });

function blockPair(
  display: string,
  oldb: Block,
  newb: Block,
  onum: number,
  nnum: number,
): Change[] {
  if (oldb.blockType === "table" && newb.blockType === "table") {
    return tableRowChanges(display, `table ${nnum}`, oldb, newb);
  }
  if (oldb.blockType === "table") {
    return [...tableRowChanges(display, `table ${onum}`, oldb, emptyTable()), blockAdded(display, newb, nnum)];
  }
  if (newb.blockType === "table") {
    return [blockRemoved(display, oldb, onum), ...tableRowChanges(display, `table ${nnum}`, emptyTable(), newb)];
  }
  const c = paraChange(display, oldb, newb, nnum);
  return c ? [c] : [];
}

export function diffDocx(old: Docx, new_: Docx, display = ""): Change[] {
  const changes: Change[] = [];
  const opcodes = sequenceMatcher(old.blocks.map(blockToken), new_.blocks.map(blockToken));
  let onum = 0;
  let nnum = 0;
  for (const [tag, i1, i2, j1, j2] of opcodes) {
    if (tag === "equal") {
      onum += i2 - i1;
      nnum += j2 - j1;
      continue;
    }
    const nOld = i2 - i1;
    const nNew = j2 - j1;
    if (tag === "replace" && nOld === nNew) {
      for (let k = 0; k < nOld; k++) {
        for (const c of blockPair(display, old.blocks[i1 + k], new_.blocks[j1 + k], onum + k + 1, nnum + k + 1)) {
          changes.push(c);
        }
      }
    } else if (tag === "replace" && nOld === 1 && nNew > 1) {
      // a split paragraph reads as one edit + extras, not a removal dump
      const first = new_.blocks[j1];
      const oldb = old.blocks[i1];
      if (first.blockType === "paragraph" && oldb.blockType === "paragraph") {
        const c = paraChange(display, oldb, first, nnum + 1);
        if (c) {
          changes.push({ ...c, note: (c.note ? c.note + "; " : "") + "split into paragraphs" });
        }
      } else {
        for (const c of blockPair(display, oldb, first, onum + 1, nnum + 1)) changes.push(c);
      }
      for (let k = 1; k < nNew; k++) changes.push(blockAdded(display, new_.blocks[j1 + k], nnum + k + 1));
    } else {
      for (let k = 0; k < nOld; k++) changes.push(blockRemoved(display, old.blocks[i1 + k], onum + k + 1));
      for (let k = 0; k < nNew; k++) changes.push(blockAdded(display, new_.blocks[j1 + k], nnum + k + 1));
    }
    onum += nOld;
    nnum += nNew;
  }
  return changes;
}
