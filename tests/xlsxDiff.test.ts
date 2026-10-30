/**
 * Golden parity: the Python engine generated tests/golden/*.json from the
 * exact fixture bytes mirrored here; the TS differ must reproduce them
 * field for field. Plus synthetic pairs for cases no engine fixture covers
 * in single-file form (sheet rename/remove).
 */

import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";

import { diffXlsx } from "../src/core/xlsxDiff";
import { extractXlsx } from "../src/extract/xlsx";

const here = import.meta.dirname;
const demo = (f: string) => readFileSync(resolve(here, "../public/demo", f));
const fix = (f: string) => readFileSync(resolve(here, "fixtures", f));
const golden = (f: string) => JSON.parse(readFileSync(resolve(here, "golden", f), "utf8"));

async function runPair(oldB: Buffer, newB: Buffer, display: string) {
  return diffXlsx(await extractXlsx(oldB), await extractXlsx(newB), display);
}

async function wbWith(sheets: [string, Record<string, string | number>][]) {
  const wb = new ExcelJS.Workbook();
  for (const [name, cells] of sheets) {
    const ws = wb.addWorksheet(name);
    for (const [addr, v] of Object.entries(cells)) ws.getCell(addr).value = v;
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

describe("golden parity with the Python engine", () => {
  it("Q3 demo pair: 14 changes, field for field", async () => {
    const got = await runPair(demo("Q3-forecast.xlsx"), demo("Q3-forecast-v2.xlsx"), "Q3-forecast.xlsx");
    expect(got).toEqual(golden("q3-forecast.json"));
  });

  it("identical file diffs to nothing", async () => {
    const got = await runPair(demo("Q3-forecast.xlsx"), demo("Q3-forecast.xlsx"), "Q3-forecast.xlsx");
    expect(got).toEqual(golden("q3-identical.json"));
    expect(got).toEqual([]);
  });

  it("formula-notes pair: all four note rules + bool guard", async () => {
    const got = await runPair(fix("formula-notes.xlsx"), fix("formula-notes-v2.xlsx"), "formula-notes.xlsx");
    expect(got).toEqual(golden("formula-notes.json"));
  });
});

describe("sheet-level alignment", () => {
  it("identical-content rename -> sheet_renamed (moved), zero cell noise", async () => {
    const oldB = await wbWith([["Budget", { A1: 1, B2: "plan" }]]);
    const newB = await wbWith([["Budget FY27", { A1: 1, B2: "plan" }]]);
    expect(await runPair(oldB, newB, "b.xlsx")).toEqual([
      {
        path: "b.xlsx",
        kind: "sheet_renamed",
        where: "Budget FY27",
        before: "Budget",
        after: "Budget FY27",
        confidence: "moved",
        note: null,
      },
    ]);
  });

  it("unrelated sheets stay honest: removed + added, never a fake rename", async () => {
    const oldB = await wbWith([["X", { A1: "apple" }]]);
    const newB = await wbWith([["Y", { A1: "banana" }]]);
    expect(await runPair(oldB, newB, "b.xlsx")).toEqual([
      { path: "b.xlsx", kind: "sheet_removed", where: "X", before: null, after: null, confidence: "exact", note: null },
      { path: "b.xlsx", kind: "sheet_added", where: "Y", before: null, after: null, confidence: "exact", note: null },
    ]);
  });
});
