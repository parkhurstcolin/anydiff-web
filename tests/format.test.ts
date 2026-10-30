import { describe, expect, it } from "vitest";

import { fmtValue, visible } from "../src/core/format";

describe("fmtValue()", () => {
  it("formats the demo line values", () => {
    expect(fmtValue(4_200_000)).toBe("4.2M");
    expect(fmtValue(4_800_000)).toBe("4.8M");
  });

  it("pins the scale boundaries (honest rounding beats promotion)", () => {
    expect(fmtValue(9999)).toBe("9999"); // below 1e4 stays raw
    expect(fmtValue(10_000)).toBe("10K"); // ".0" stripped
    expect(fmtValue(999_999)).toBe("1000K"); // NOT "1M"
    expect(fmtValue(1_000_000)).toBe("1M");
    expect(fmtValue(999_999_999)).toBe("1000M");
    expect(fmtValue(-4_800_000_000)).toBe("-4.8B");
  });

  it("matches the engine's non-numeric spellings", () => {
    expect(fmtValue(null)).toBe("");
    expect(fmtValue(true)).toBe("TRUE");
    expect(fmtValue(false)).toBe("FALSE");
    expect(fmtValue("4.2M")).toBe("4.2M"); // strings pass through untouched
    expect(fmtValue(42)).toBe("42");
    expect(fmtValue(3.5)).toBe("3.5");
  });

  it("renders dates, with time only when present", () => {
    expect(fmtValue(new Date(2026, 9, 30))).toBe("2026-10-30");
    expect(fmtValue(new Date(2026, 9, 30, 14, 3, 9))).toBe("2026-10-30 14:03:09");
  });
});

describe("visible()", () => {
  it("leaves ordinary text byte-identical", () => {
    const t = "Sheet2!B14: 4.2M -> 4.8M";
    expect(visible(t)).toBe(t);
  });

  it("escapes the spoofers the engine escaped (issue 31)", () => {
    expect(visible("invoice\u202egpj.exe")).toBe("invoice\\u202egpj.exe"); // RLO
    expect(visible("\x1b]8;;http")).toBe("\\x1b]8;;http"); // ESC: no terminal OSC
    expect(visible("a\u200bb")).toBe("a\\u200bb"); // ZWSP (Cf)
  });

  it("keeps combining marks and simple escapes readable", () => {
    expect(visible("é")).toBe("é"); // combining acute stays visible
    expect(visible("a\nb")).toBe("a\\nb");
  });

  it("is idempotent — renderers may compose outputs freely", () => {
    const once = visible("in\u202evo\x1bce");
    expect(visible(once)).toBe(once);
  });
});
