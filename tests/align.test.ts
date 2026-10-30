import { describe, expect, it } from "vitest";

import { alignByAnchors, jaccard, keyCmp } from "../src/core/align";

const t = (...s: string[]) => new Set(s);
const m = <K extends string | number>(entries: [K, ReadonlySet<string> | (() => ReadonlySet<string>)][]) =>
  new Map(entries);

describe("jaccard", () => {
  it("empty vs empty is a match; disjoint is zero", () => {
    expect(jaccard(t(), t())).toBe(1.0);
    expect(jaccard(t("a"), t("b"))).toBe(0);
    expect(jaccard(t("a", "b", "c"), t("b", "c", "d"))).toBeCloseTo(0.5, 10);
  });
});

describe("keyCmp", () => {
  it("orders numbers numerically, strings by code point", () => {
    expect([10, 9].sort(keyCmp)).toEqual([9, 10]);
    expect(["a", "B"].sort(keyCmp)).toEqual(["B", "a"]); // code point, not locale
  });
});

describe("alignByAnchors — sheets (exact keys anchored)", () => {
  it("all names shared -> exact pairs, zero thunks executed", () => {
    let built = 0;
    const lazy = (set: Set<string>) => () => {
      built++;
      return set;
    };
    const a = alignByAnchors(
      m([["A", lazy(t("x"))], ["B", lazy(t("y"))]]),
      m([["A", lazy(t("x"))], ["B", lazy(t("y2"))]]),
    );
    expect(a.pairs).toEqual([
      { old: "A", new: "A", confidence: "exact" },
      { old: "B", new: "B", confidence: "exact" },
    ]);
    expect(built).toBe(0); // day-6 perf rule: exact matches never tokenize
  });

  it("identical-content rename -> one moved pair, nothing removed/added", () => {
    const a = alignByAnchors(m([["Old", t("1", "2", "3")]]), m([["New", t("1", "2", "3")]]));
    expect(a.pairs).toEqual([{ old: "Old", new: "New", confidence: "moved" }]);
    expect(a.removed).toEqual([]);
    expect(a.added).toEqual([]);
  });

  it("heavily-edited rename (sim 0.6) degrades to manual-inspect", () => {
    const a = alignByAnchors(
      m([["Old", t("1", "2", "3", "4", "5", "6", "7", "8")]]),
      m([["New", t("1", "2", "3", "4", "5", "6", "9", "10")]]),
    );
    expect(a.pairs).toEqual([{ old: "Old", new: "New", confidence: "manual-inspect" }]);
  });

  it("unrelated sheets never pair — removed + added honestly", () => {
    const a = alignByAnchors(m([["Gone", t("1")]]), m([["Fresh", t("2")]]));
    expect(a.pairs).toEqual([]);
    expect(a.removed).toEqual(["Gone"]);
    expect(a.added).toEqual(["Fresh"]);
  });

  it("contested match (two near-equal rivals) -> manual-inspect, loser added", () => {
    const a = alignByAnchors(
      m([["X", t("1", "2", "3")]]),
      m([["Y", t("1", "2", "3")], ["Z", t("1", "2", "3")]]),
    );
    expect(a.pairs).toEqual([{ old: "X", new: "Y", confidence: "manual-inspect" }]);
    expect(a.added).toEqual(["Z"]);
  });
});

describe("alignByAnchors — slides (no exact anchoring)", () => {
  it("same-number-swap pairs by CONTENT, not number", () => {
    const a = alignByAnchors(
      m([[1, t("intro")], [2, t("chart")]]),
      m([[1, t("chart")], [2, t("intro")]]),
      false,
    );
    expect(a.pairs).toEqual([
      { old: 1, new: 2, confidence: "moved" },
      { old: 2, new: 1, confidence: "moved" },
    ]);
    expect(a.removed).toEqual([]);
    expect(a.added).toEqual([]);
  });

  it("pairs sort by numeric key order (string sort would say 10 < 9)", () => {
    const a = alignByAnchors(
      m([[9, t("a")], [10, t("b")]]),
      m([[9, t("b")], [10, t("a")]]),
      false,
    );
    expect(a.pairs.map((p) => p.old)).toEqual([9, 10]);
  });

  it("every key lands in exactly one bucket", () => {
    const a = alignByAnchors(
      m([[1, t("x")], [2, t("y")], [3, t("z")]]),
      m([[1, t("x2")], [4, t("y")]]),
      false,
    );
    const seen = new Set([...a.pairs.map((p) => p.old), ...a.removed]);
    expect(seen.size).toBe(3); // 1 pairs-or-? , 2 paired to 4, 3 removed
    expect(new Set([...a.pairs.map((p) => p.new), ...a.added]).size).toBe(2);
  });
});
