/**
 * XLSX differ — port of the engine's diff.py (_values_differ, _addr_key,
 * _cell_changes, diff_xlsx) against align.ts + model.ts. The golden tests
 * (tests/golden/*.json, produced by the Python engine) are the referee.
 */

import { alignByAnchors, keyCmp } from "./align";
import { foldCell, newSumXor, sameSumXor } from "./hash";
import { change, type Change, type ChangeValue } from "./model";
import type { CellValue, Sheet } from "../extract/xlsx";

// note wording must match diff.py verbatim: only the differ knows whether
// "value unchanged" was actually verified against cached values
export const NOTE_VALUE_UNCHANGED = "formula changed, value unchanged";
export const NOTE_VALUES_UNVERIFIED = "formula changed, cached values unavailable";
export const NOTE_FORMULA_CHANGED = "formula changed";
export const NOTE_FORMULA_ALSO = "formula also changed";

const pad = (n: number) => String(n).padStart(2, "0");

/** Excel dates are naive, like Python's; str() of a date/datetime is the
 *  JSON-boundary spelling (Date itself is outside the Change contract). */
function pyStr(d: Date): string {
  const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  return d.getHours() || d.getMinutes() || d.getSeconds()
    ? `${date} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
    : date;
}

function toContract(v: CellValue): ChangeValue {
  return v instanceof Date ? pyStr(v) : v;
}

function valuesDiffer(b: CellValue, a: CellValue): boolean {
  // Excel TRUE/FALSE are booleans; JS true == 1 would hide 1 -> TRUE edits,
  // so a bool-vs-non-bool pair is always a change (engine's guard).
  if ((typeof b === "boolean") !== (typeof a === "boolean")) return true;
  if (b instanceof Date || a instanceof Date) {
    if (!(b instanceof Date && a instanceof Date)) return true;
    return b.getTime() !== a.getTime();
  }
  return b !== a;
}

/** (row, col) — addresses sort by row first, then column. */
function addrKey(addr: string): [number, number] {
  const m = /^([A-Z]+)(\d+)$/.exec(addr);
  if (!m) return [1e9, 0];
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]), col];
}

/** Internal alignment tokens — spelling is TS-private (never exchanged
 *  with the engine), only stability within one run matters. */
function sheetTokens(sheet: Sheet): ReadonlySet<string> {
  const out = new Set<string>();
  for (const c of sheet.cells.values()) {
    const v = c.value === null ? "None" : c.value instanceof Date ? pyStr(c.value) : JSON.stringify(c.value);
    out.add(`${c.addr}|${v}|${c.formula ?? "None"}`);
  }
  return out;
}

/** Content fingerprint of a sheet: a pair with equal fingerprints cannot
 *  contain cell changes (address, value and formula all fold in), so the
 *  differ skips cellChanges entirely — no key-union allocation, no per-cell
 *  lookups (the day-6 rule for 400k-cell sheets). */
export const sheetFingerprint = (sheet: Sheet) => {
  const acc = newSumXor();
  for (const c of sheet.cells.values()) foldCell(acc, c.addr, c.value, c.formula);
  return acc;
};

function cellChanges(
  path: string,
  name: string,
  oldSheet: Sheet,
  newSheet: Sheet,
  confidence: string,
): Change[] {
  const cellConf = confidence === "manual-inspect" ? "manual-inspect" : "exact";
  const pending: [string, Change][] = [];
  // insertion order now, sorted by (row, col) at the end — sorting only the
  // CHANGED addresses (engine's day-6 perf rule for 400k-cell sheets)
  const addrs = [...oldSheet.cells.keys()];
  for (const k of newSheet.cells.keys()) if (!oldSheet.cells.has(k)) addrs.push(k);
  for (const addr of addrs) {
    const before = oldSheet.cells.get(addr);
    const after = newSheet.cells.get(addr);
    const bVal = before?.value ?? null;
    const aVal = after?.value ?? null;
    if (before && after && before.formula === after.formula && !valuesDiffer(bVal, aVal)) {
      continue; // identical cell: the overwhelmingly common case
    }
    const formulasDiffer = !!(before && after && before.formula !== after.formula);
    const bothFormulas = !!(before && after && before.formula && after.formula);
    // a formula cell with no cached value has an UNKNOWN value, not a null
    // one; its apparent (un)equal value can never verify a claim
    const bUnknown = !!(before?.formula && bVal === null);
    const aUnknown = !!(after?.formula && aVal === null);
    let note: string | null = null;
    let bDisp: CellValue;
    let aDisp: CellValue;
    if (valuesDiffer(bVal, aVal)) {
      if (bothFormulas && formulasDiffer && !(bUnknown || aUnknown)) note = NOTE_FORMULA_ALSO;
      else if (formulasDiffer && (bUnknown || aUnknown)) note = NOTE_VALUES_UNVERIFIED;
      bDisp = before && bUnknown ? before.formula : bVal;
      aDisp = after && aUnknown ? after.formula : aVal;
    } else if (formulasDiffer) {
      if (bothFormulas && !(bUnknown || aUnknown)) note = NOTE_VALUE_UNCHANGED;
      else if (bothFormulas) note = NOTE_VALUES_UNVERIFIED; // e.g. openpyxl-written files: no cached values
      else note = NOTE_FORMULA_CHANGED;
      // literals keep their value so a literal -> formula edit doesn't
      // render as blank -> formula
      bDisp = before?.formula ?? bVal;
      aDisp = after?.formula ?? aVal;
    } else {
      bDisp = bVal;
      aDisp = aVal;
    }
    pending.push([
      addr,
      change({
        path,
        kind: "cell",
        where: `${name}!${addr}`,
        before: toContract(bDisp),
        after: toContract(aDisp),
        confidence: cellConf,
        note,
      }),
    ]);
  }
  pending.sort((x, y) => {
    const [ra, ca] = addrKey(x[0]);
    const [rb, cb] = addrKey(y[0]);
    return ra - rb || ca - cb;
  });
  return pending.map(([, c]) => c);
}

export function diffXlsx(
  oldSheets: Map<string, Sheet>,
  newSheets: Map<string, Sheet>,
  display = "",
): Change[] {
  const asTokens = (m: Map<string, Sheet>) =>
    new Map<string, () => ReadonlySet<string>>(
      [...m].map(([n, s]) => [n, () => sheetTokens(s)]), // lazy: exact-name pairs never tokenize
    );
  const alignment = alignByAnchors(asTokens(oldSheets), asTokens(newSheets));
  const fpOld = new Map([...oldSheets].map(([n, s]) => [n, sheetFingerprint(s)]));
  const fpNew = new Map([...newSheets].map(([n, s]) => [n, sheetFingerprint(s)]));
  const changes: Change[] = [];
  for (const pair of alignment.pairs) {
    if (pair.old !== pair.new) {
      changes.push(
        change({
          path: display,
          kind: "sheet_renamed",
          where: pair.new,
          before: pair.old,
          after: pair.new,
          confidence: pair.confidence,
        }),
      );
    }
    if (sameSumXor(fpOld.get(pair.old)!, fpNew.get(pair.new)!)) continue; // instant: no cell can differ
    for (const c of cellChanges(
      display,
      pair.new,
      oldSheets.get(pair.old)!,
      newSheets.get(pair.new)!,
      pair.confidence,
    )) {
      changes.push(c);
    }
  }
  for (const name of [...alignment.removed].sort(keyCmp)) {
    changes.push(change({ path: display, kind: "sheet_removed", where: name, before: null, after: null }));
  }
  for (const name of [...alignment.added].sort(keyCmp)) {
    changes.push(change({ path: display, kind: "sheet_added", where: name, before: null, after: null }));
  }
  return changes;
}
