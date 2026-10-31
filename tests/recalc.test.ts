/**
 * Recalculation is an opt-in input transformer: it swaps saved values for
 * computed ones BEFORE the differ runs. The differ and the golden parity
 * never know it exists — these tests cover the swap and its honesty rules.
 */

import { describe, expect, it } from "vitest";

import { recalcSheets } from "../src/core/recalc";
import type { Cell, Sheet } from "../src/extract/xlsx";

const sheet = (name: string, entries: [string, Cell["value"], string | null][]): Sheet => ({
  name,
  cells: new Map(entries.map(([addr, value, formula]) => [addr, { addr, value, formula }])),
});
const val = (s: Sheet, addr: string) => s.cells.get(addr)?.value;

describe("recalcSheets", () => {
  it("computes formulas, keeps formula text and untouched values", () => {
    const input = new Map([
      [
        "S1",
        sheet("S1", [
          ["A1", 2, null],
          ["B1", 3, null],
          ["C1", 0, "=SUM(A1:B1)"],
          ["D1", "text", null],
        ]),
      ],
    ]);
    const r = recalcSheets(input);
    expect(r.skipped).toBeNull();
    expect(r.computed).toBe(1);
    expect(r.failed).toBe(0);
    expect(val(r.sheets.get("S1")!, "C1")).toBe(5);
    expect(r.sheets.get("S1")!.cells.get("C1")!.formula).toBe("=SUM(A1:B1)");
    expect(val(r.sheets.get("S1")!, "D1")).toBe("text");
    // input must not be mutated — the as-saved view stays available
    expect(val(input.get("S1")!, "C1")).toBe(0);
  });

  it("rescues valueless formulas from files that never calculated", () => {
    const input = new Map([
      [
        "S",
        sheet("S", [
          ["A1", 10, null],
          ["A2", 32, null],
          ["A3", null, "=SUM(A1:A2)"],
        ]),
      ],
    ]);
    const r = recalcSheets(input);
    expect(val(r.sheets.get("S")!, "A3")).toBe(42);
    expect(r.failed).toBe(0);
  });

  it("follows cross-sheet references", () => {
    const input = new Map([
      ["S1", sheet("S1", [["A1", null, "='S2'!B2*10"]])],
      ["S2", sheet("S2", [["B2", 4.5, null]])],
    ]);
    expect(val(recalcSheets(input).sheets.get("S1")!, "A1")).toBe(45);
  });

  it("engine errors and date formulas keep the saved value and count as failed", () => {
    const input = new Map([
      [
        "S",
        sheet("S", [
          ["A1", 99, "=1/0"],
          ["B1", 7, "=NOSUCHFN(1)"],
          ["C1", new Date(2026, 0, 15), "=TODAY()"],
        ]),
      ],
    ]);
    const r = recalcSheets(input);
    expect(r.computed).toBe(0);
    expect(r.failed).toBe(3);
    expect(val(r.sheets.get("S")!, "A1")).toBe(99);
    expect(val(r.sheets.get("S")!, "B1")).toBe(7);
    expect(val(r.sheets.get("S")!, "C1")).toBeInstanceOf(Date);
  });

  it("skips absurdly large grids instead of hanging", () => {
    const input = new Map([["S", sheet("S", [["A1", 1, null], ["ZZ5000", 2, "=A1+1"]])]]);
    const r = recalcSheets(input);
    expect(r.skipped).toMatch(/too large/);
    expect(r.sheets).toBe(input);
  });

  it("a stale cached value becomes a real diff after recalculation", () => {
    // the classic: upstream number edited by a tool that never re-evaluates
    const oldWb = new Map([
      [
        "S",
        sheet("S", [
          ["A1", 1, null],
          ["B1", 3, "=A1+2"],
        ]),
      ],
    ]);
    const newWb = new Map([
      [
        "S",
        sheet("S", [
          ["A1", 5, null],
          ["B1", 3, "=A1+2"], // cached value left stale by the writer
        ]),
      ],
    ]);
    const oldR = recalcSheets(oldWb);
    const newR = recalcSheets(newWb);
    expect(val(newR.sheets.get("S")!, "B1")).toBe(7);
    expect(oldR.failed + newR.failed).toBe(0);
  });
});
