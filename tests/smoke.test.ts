import { readFileSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

// fileURLToPath, not .pathname: this workspace path contains a space
const DEMO = fileURLToPath(new URL("../public/demo/", import.meta.url));

// Smoke: the demo fixtures copied from the engine repo are real, non-empty,
// and their zip magic bytes are intact — the golden-oracle phase builds on this.
describe("demo fixtures", () => {
  it("copies every demo file non-empty", () => {
    for (const name of readdirSync(DEMO)) {
      if (name.endsWith(".md")) continue;
      expect(readFileSync(DEMO + name).byteLength, name).toBeGreaterThan(1000);
    }
  });

  it("xlsx/docx files carry zip magic (PK)", () => {
    for (const name of readdirSync(DEMO)) {
      if (!/\.(xlsx|docx)$/.test(name)) continue;
      const head = readFileSync(DEMO + name).subarray(0, 2).toString("latin1");
      expect(head, name).toBe("PK");
    }
  });
});
