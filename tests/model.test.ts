import { describe, expect, it } from "vitest";

import { change, isChange, type Change } from "../src/core/model";

const base = {
  path: "Q3-forecast.xlsx",
  kind: "cell" as const,
  where: "Sheet2!B14",
  before: "4.2M",
  after: "4.8M",
};

describe("change()", () => {
  it("applies the Python dataclass defaults", () => {
    const c: Change = change(base);
    expect(c.confidence).toBe("exact");
    expect(c.note).toBeNull();
  });

  it("mirrors __post_init__: unknown kind throws", () => {
    expect(() => change({ ...base, kind: "paragraphx" as never })).toThrow(
      /unknown change kind|invalid change record/,
    );
  });

  it("mirrors __post_init__: unknown confidence throws", () => {
    expect(() => change({ ...base, confidence: "probably" as never })).toThrow(
      /invalid change record/,
    );
  });
});

describe("isChange()", () => {
  it("accepts a full record", () => {
    expect(isChange({ ...base, confidence: "moved", note: "x" })).toBe(true);
  });

  it("rejects junk the same way the engine would", () => {
    expect(isChange(null)).toBe(false);
    expect(isChange({ ...base, kind: "slide " })).toBe(false); // trailing space is a different kind
    expect(isChange({ ...base, confidence: "probably" })).toBe(false);
    expect(isChange({ ...base, before: { nested: true } })).toBe(false);
  });
});
