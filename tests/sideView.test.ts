/**
 * Side-by-side view model: the line-up must MIRROR the differ — every cell
 * change lands as a changed grid cell in its sheet pane, and docx lines
 * reuse the differ's opcodes (same / edit / added / removed, table pairing).
 * A highlight contradicting the change list is the failure mode to catch.
 */

// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

import { diffDocx } from "../src/core/docxDiff";
import { diffXlsx } from "../src/core/xlsxDiff";
import { sideDocx, sideXlsx } from "../src/core/sideView";
import { extractDocx } from "../src/extract/docx";
import type { ParagraphBlock, TableBlock } from "../src/extract/docx";
import type { Sheet } from "../src/extract/xlsx";
import { extractXlsx } from "../src/extract/xlsx";

const here = import.meta.dirname;
const demo = (f: string) => readFileSync(resolve(here, "../public/demo", f));

const sheet = (name: string, entries: [string, number | string][]): Sheet => ({
  name,
  cells: new Map(entries.map(([addr, v]) => [addr, { addr, value: v, formula: null }])),
});
const para = (text: string, style = "Normal"): ParagraphBlock => ({
  blockType: "paragraph",
  style,
  runs: [{ text, bold: false, italic: false }],
});
const table = (rows: string[][]): TableBlock => ({ blockType: "table", rows });

describe("sideXlsx", () => {
  it("Q3 demo: every cell change lands on a highlighted grid cell", async () => {
    const oldS = await extractXlsx(demo("Q3-forecast.xlsx"));
    const newS = await extractXlsx(demo("Q3-forecast-v2.xlsx"));
    const changes = diffXlsx(oldS, newS, "v2");
    const panes = sideXlsx(changes, oldS, newS);
    const cellChanges = changes.filter((c) => c.kind === "cell");
    expect(cellChanges.length).toBeGreaterThan(0);
    const sides = cellChanges.reduce((n, c) => n + (c.before !== null ? 1 : 0) + (c.after !== null ? 1 : 0), 0);
    const marked = panes.reduce(
      (n, p) => n + [...p.oldRows, ...p.newRows].reduce((m, r) => m + r.cells.filter((c) => c?.changed).length, 0),
      0,
    );
    expect(marked).toBe(sides);
    expect(panes.length).toBe(newS.size + changes.filter((c) => c.kind === "sheet_removed").length);
  });

  it("window frames the change with context rows and columns", () => {
    const o = new Map([["S", sheet("S", [["A1", "head"], ["F10", 1]])]]);
    const n = new Map([["S", sheet("S", [["A1", "head"], ["F10", 2]])]]);
    const pane = sideXlsx(diffXlsx(o, n, "wb"), o, n)[0];
    expect(pane.row0).toBe(8);
    expect(pane.col0).toBe(4);
    expect(pane.oldRows.length).toBe(5); // rows 8..12
    expect(pane.oldRows[0].cells.length).toBe(5); // cols D..H
    const marked = [...pane.oldRows, ...pane.newRows].flatMap((r) => r.cells.filter((c) => c?.changed)).length;
    expect(marked).toBe(2); // 1 and 2, one per side
  });

  it("renamed sheets pair into one pane with the rename label", () => {
    const o = new Map([
      ["Sales", sheet("Sales", [["A1", 1]])],
      ["Notes", sheet("Notes", [["A1", "x"]])],
    ]);
    const n = new Map([
      ["Revenue", sheet("Revenue", [["A1", 1]])],
      ["Notes", sheet("Notes", [["A1", "x"]])],
    ]);
    const changes = diffXlsx(o, n, "wb");
    expect(changes.some((c) => c.kind === "sheet_renamed")).toBe(true);
    const pane = sideXlsx(changes, o, n).find((p) => p.rename);
    expect(pane?.rename).toEqual({ old: "Sales", new: "Revenue" });
    expect(pane?.oldRows[0].cells[0]?.text).toBe("1");
  });

  it("removed and added sheets get one-sided panes", () => {
    const o = new Map([
      ["Keep", sheet("Keep", [["A1", 1]])],
      ["Gone", sheet("Gone", [["A1", 1]])],
    ]);
    const n = new Map([
      ["Keep", sheet("Keep", [["A1", 1]])],
      ["Fresh", sheet("Fresh", [["A1", 2]])],
    ]);
    const panes = sideXlsx(diffXlsx(o, n, "wb"), o, n);
    expect(panes.find((p) => p.kind === "removed")?.note).toBe("sheet removed");
    const added = panes.find((p) => p.kind === "added");
    expect(added?.note).toBe("sheet added");
    expect(added?.newRows.length).toBeGreaterThan(0);
    expect(added?.oldRows.length).toBe(0);
  });
});
describe("sideDocx", () => {
  it("meeting-notes demo: heading edit, context lines and an inserted table row", async () => {
    const a = await extractDocx(demo("meeting-notes.docx"));
    const b = await extractDocx(demo("meeting-notes-v2.docx"));
    expect(diffDocx(a, b, "v2").length).toBeGreaterThan(0);
    const lines = sideDocx(a, b);
    expect(lines.some((l) => l.type === "same")).toBe(true);
    const edit = lines.find(
      (l) => l.type === "edit" && (l.styleChanged || l.oldText !== l.newText),
    );
    expect(edit).toBeDefined();
    const tables = lines.filter((l) => l.type === "tablePair");
    expect(tables.length).toBeGreaterThan(0);
    expect(tables.some((l) => l.type === "tablePair" && l.rows.some((r) => r.status === "added" && r.newCells))).toBe(true);
  });

  it("context, additions, removals and paragraph edits map to line kinds", () => {
    const a = { blocks: [para("intro"), para("drop me")] };
    const b = { blocks: [para("intro"), para("polish this now"), para("fresh point")] };
    const kinds = sideDocx(a, b).map((l) => l.type);
    expect(kinds).toEqual(["same", "edit", "added"]);
  });

  it("style-only edit is an edit line flagged styleChanged", () => {
    const a = { blocks: [para("heading", "Normal")] };
    const b = { blocks: [para("heading", "Heading1")] };
    const [line] = sideDocx(a, b);
    expect(line.type).toBe("edit");
    expect(line.type === "edit" && line.styleChanged).toBe(true);
    expect(line.type === "edit" && line.oldText === line.newText).toBe(true);
  });

  it("equal-count table replace pairs rows as edits; inserts stay one-sided", () => {
    const a = { blocks: [table([["h", "x"], ["keep", "1"]])] };
    const b = { blocks: [table([["h", "y"], ["keep", "1"], ["extra", "2"]])] };
    const line = sideDocx(a, b)[0];
    expect(line.type).toBe("tablePair");
    if (line.type !== "tablePair") return;
    expect(line.rows[0].status).toBe("edit");
    expect(line.rows[1].status).toBe("same");
    expect(line.rows[2].status).toBe("added");
    expect(line.rows[2].oldCells).toBe(null);
  });

  it("table replaced by paragraph dumps the table as a table + added prose", () => {
    const a = { blocks: [table([["a", "b"]])] };
    const b = { blocks: [para("prose instead")] };
    const lines = sideDocx(a, b);
    expect(lines.map((l) => l.type)).toEqual(["tableDump", "added"]);
    expect(lines[0].type === "tableDump" && lines[0].side).toBe("old");
  });
});

describe("render smoke (catches broken JSX paths before the browser does)", () => {
  it("SideBySideView renders demo docx and xlsx pairs to markup", async () => {
    const { renderToStaticMarkup } = await import("react-dom/server");
    const { default: SideBySideView } = await import("../src/components/SideBySideView");
    const { createElement } = await import("react");

    const d1 = await extractDocx(demo("meeting-notes.docx"));
    const d2 = await extractDocx(demo("meeting-notes-v2.docx"));
    const html = renderToStaticMarkup(
      createElement(SideBySideView, {
        pair: { kind: "docx", old: d1, new: d2 },
        changes: diffDocx(d1, d2, "v2"),
        oldName: "meeting-notes.docx",
        newName: "meeting-notes-v2.docx",
      }),
    );
    expect(html).toContain("before · meeting-notes.docx");
    expect(html).toContain("bg-red-"); // at least one removed-side highlight

    const x1 = await extractXlsx(demo("Q3-forecast.xlsx"));
    const x2 = await extractXlsx(demo("Q3-forecast-v2.xlsx"));
    const html2 = renderToStaticMarkup(
      createElement(SideBySideView, {
        pair: { kind: "xlsx", old: x1, new: x2 },
        changes: diffXlsx(x1, x2, "v2"),
      }),
    );
    expect(html2).toContain("bg-red-100");
    expect(html2).toContain("bg-emerald-100");
  });
});
