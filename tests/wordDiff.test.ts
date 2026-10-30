import { describe, expect, it } from "vitest";

import { diffWords, tokenize } from "../src/core/wordDiff";

const flat = (s: string) => diffWords(s, s);
const rebuild = (segs: { text: string; kind: string }[], kinds: string[]) =>
  segs.filter((s) => kinds.includes(s.kind)).map((s) => s.text).join("");

describe("diffWords", () => {
  it("identical text is one same-segment", () => {
    expect(flat("hello world")).toEqual([{ text: "hello world", kind: "same" }]);
  });

  it("segments always rebuild both sides exactly, in order", () => {
    const [before, after] = ["Q3 revenue plan v2 draft", "Q3 revenue plan v3 final"];
    const segs = diffWords(before, after);
    expect(rebuild(segs, ["same", "removed"])).toBe(before);
    expect(rebuild(segs, ["same", "added"])).toBe(after);
  });

  it("marks only the changed word (word granularity: FY26 is one token)", () => {
    expect(diffWords("revenue plan", "revenue plan")).toHaveLength(1);
    const segs = diffWords("Contingency budget FY26", "Contingency budget FY27");
    expect(segs).toEqual([
      { text: "Contingency budget ", kind: "same" },
      { text: "FY26", kind: "removed" },
      { text: "FY27", kind: "added" },
    ]);
  });

  it("empty side is a whole-side mark", () => {
    expect(diffWords("", "new")).toEqual([{ text: "new", kind: "added" }]);
    expect(diffWords("gone", "")).toEqual([{ text: "gone", kind: "removed" }]);
  });

  it("pathological length falls back instead of stalling", () => {
    const big = "w ".repeat(500).trim();
    const segs = diffWords(big, big + " tail");
    expect(segs.some((s) => s.kind === "removed" && s.text.length > 400)).toBe(true);
  });

  it("tokenize keeps whitespace so joins round-trip", () => {
    expect(tokenize("a  b ").join("")).toBe("a  b ");
  });
});
