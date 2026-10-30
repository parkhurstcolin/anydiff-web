/**
 * Fingerprint rules: equal => provably-equal content (within birthday risk);
 * unequal decides nothing and never short-circuits; order matters where the
 * differ cares about order (docx blocks) and nowhere it doesn't (sheet cells).
 */

import { describe, expect, it } from "vitest";

import { cyrb53, foldBlock, foldCell, hashBytes, newSumXor, sameSumXor } from "../src/core/hash";
import { diffDocx, docxFingerprint } from "../src/core/docxDiff";
import { diffXlsx, sheetFingerprint } from "../src/core/xlsxDiff";
import type { ParagraphBlock, TableBlock } from "../src/extract/docx";
import type { Cell, Sheet } from "../src/extract/xlsx";

const sheet = (name: string, entries: [string, Cell["value"], string | null][]): Sheet => ({
  name,
  cells: new Map(entries.map(([addr, value, formula]) => [addr, { addr, value, formula }])),
});
const para = (text: string, style = "Normal", bold = false): ParagraphBlock => ({
  blockType: "paragraph",
  style,
  runs: [{ text, bold, italic: false }],
});
const table = (rows: string[][]): TableBlock => ({ blockType: "table", rows });

describe("cyrb53 / hashBytes", () => {
  it("is deterministic and separates near-identical inputs", () => {
    expect(cyrb53("anything")).toBe(cyrb53("anything"));
    expect(cyrb53("a")).not.toBe(cyrb53("b"));
    expect(cyrb53("abc")).not.toBe(cyrb53("abd"));
    expect(cyrb53("")).not.toBe(cyrb53("\0"));
  });

  it("fingerprints bytes by content, order and length", () => {
    const u = new Uint8Array([1, 2, 3, 4, 5]);
    expect(hashBytes(u)).toBe(hashBytes(new Uint8Array([1, 2, 3, 4, 5])));
    const flip = new Uint8Array(u);
    flip[3] ^= 1;
    expect(hashBytes(flip)).not.toBe(hashBytes(u));
    expect(hashBytes(new Uint8Array([5, 4, 3, 2, 1]))).not.toBe(hashBytes(u));
    expect(hashBytes(new Uint8Array([1, 2, 3, 4, 5, 0]))).not.toBe(hashBytes(u));
  });
});

describe("sheetFingerprint", () => {
  it("ignores cell insertion order (a sheet has no cell order)", () => {
    const a = sheet("S", [["A1", 1, null], ["B2", "x", "=SUM(A1:A2)"]]);
    const b = sheet("S", [["B2", "x", "=SUM(A1:A2)"], ["A1", 1, null]]);
    expect(sameSumXor(sheetFingerprint(a), sheetFingerprint(b))).toBe(true);
  });

  it("every field the differ reads moves the fingerprint", () => {
    const fp = (entries: [string, Cell["value"], string | null][]) => sheetFingerprint(sheet("S", entries));
    const base = fp([["A1", 1, null], ["B2", 2, "=A1"]]);
    expect(sameSumXor(base, fp([["A1", 1.5, null], ["B2", 2, "=A1"]]))).toBe(false); // value
    expect(sameSumXor(base, fp([["A1", 1, null], ["B2", 2, "=A2"]]))).toBe(false); // formula
    expect(sameSumXor(base, fp([["A1", 1, null]]))).toBe(false); // cell removed
    expect(sameSumXor(base, fp([["A1", 1, null], ["B2", 2, null]]))).toBe(false); // formula gone
    expect(sameSumXor(base, fp([["A1", true, null], ["B2", 2, "=A1"]]))).toBe(false); // TRUE vs 1
    const t = new Date(2026, 0, 1);
    expect(sameSumXor(fp([["A1", t, null]]), fp([["A1", new Date(2026, 0, 1), null]]))).toBe(true); // date by time
    expect(sameSumXor(fp([["A1", t, null]]), fp([["A1", new Date(2026, 0, 2), null]]))).toBe(false);
  });


  it("NaN never authorizes a skip — the differ still flags NaN -> NaN", () => {
    const o = new Map([["S", sheet("S", [["A1", NaN, null]])]]);
    const n = new Map([["S", sheet("S", [["A1", NaN, null]])]]);
    expect(sameSumXor(sheetFingerprint(o.get("S")!), sheetFingerprint(n.get("S")!))).toBe(false);
    expect(diffXlsx(o, n, "wb")).toHaveLength(1); // engine rule: NaN != NaN
  });

  it("identical sheets skip to zero changes through the differ", () => {
    const o = new Map([["S", sheet("S", [["A1", 1, null]])]]);
    const n = new Map([["S", sheet("S", [["A1", 1, null]])]]);
    expect(diffXlsx(o, n, "wb")).toHaveLength(0);
  });
});

describe("docxFingerprint", () => {
  it("identical documents skip to zero changes", () => {
    const a = { blocks: [para("title", "Heading1"), para("body"), table([["h", "v"]])] };
    const b = { blocks: [para("title", "Heading1"), para("body"), table([["h", "v"]])] };
    expect(docxFingerprint(a)).toBe(docxFingerprint(b));
    expect(diffDocx(a, b, "v2")).toHaveLength(0);
  });

  it("block order is significant — a moved paragraph is a change", () => {
    const a = { blocks: [para("first"), para("second")] };
    const b = { blocks: [para("second"), para("first")] };
    expect(docxFingerprint(a)).not.toBe(docxFingerprint(b));
    expect(diffDocx(a, b, "v2").length).toBeGreaterThan(0);
  });

  it("every field the differ reads moves the fingerprint", () => {
    const base = docxFingerprint({ blocks: [para("text"), table([["a", "b"]])] });
    const vs = (blocks: (ParagraphBlock | TableBlock)[]) => docxFingerprint({ blocks }) !== base;
    expect(vs([para("text!"), table([["a", "b"]])])).toBe(true); // run text
    expect(vs([para("text", "Heading1"), table([["a", "b"]])])).toBe(true); // style
    expect(vs([para("text", "Normal", true), table([["a", "b"]])])).toBe(true); // bold
    expect(vs([para("text"), table([["a", "B"]])])).toBe(true); // cell text
    expect(vs([para("text"), table([["a"], ["b"]])])).toBe(true); // row layout
    expect(vs([para("text")])).toBe(true); // table removed
  });

  it("run text content survives a token-blind fold: 'ab' != 'a'+'b' runs", () => {
    const one = docxFingerprint({ blocks: [{ blockType: "paragraph", style: "Normal", runs: [{ text: "ab", bold: false, italic: false }] }] });
    const two = docxFingerprint({ blocks: [{ blockType: "paragraph", style: "Normal", runs: [{ text: "a", bold: false, italic: false }, { text: "b", bold: false, italic: false }] }] });
    expect(one).not.toBe(two); // differ calls that a run-formatting change
  });
});

describe("accumulator", () => {
  it("commutative fold: insertion order cannot hide or fake differences", () => {
    const acc = newSumXor();
    const o = newSumXor();
    foldCell(acc, "A1", 1, null);
    foldCell(acc, "B1", "x", "=A1");
    foldCell(o, "B1", "x", "=A1");
    foldCell(o, "A1", 1, null);
    expect(sameSumXor(acc, o)).toBe(true);
    const b = newSumXor();
    foldBlock(b, para("one"));
    expect(sameSumXor(acc, b)).toBe(false);
  });
});
