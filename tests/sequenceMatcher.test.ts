import { describe, expect, it } from "vitest";

import { sequenceMatcher, type Opcode } from "../src/core/sequenceMatcher";

// ground truth from python3 -c "from difflib import SequenceMatcher; ...autojunk=False"
const EXPECTED: Record<string, [string, string]> = {
  abx_cd: ["abxcd", "abcd"],
  qabxcd_abycdf: ["qabxcd", "abycdf"],
  empty_both: ["", ""],
  a_to_empty: ["a", ""],
  empty_to_a: ["", "a"],
  full_equal: ["abc", "abc"],
  replace_run: ["abcd", "xycz"],
};

const expectOp = (name: string, want: Opcode[]) => {
  const [a, b] = EXPECTED[name];
  expect(sequenceMatcher([...a], [...b])).toEqual(want);
};

describe("sequenceMatcher == difflib.SequenceMatcher(autojunk=False).get_opcodes()", () => {
  it("classic abxcd/abcd", () => expectOp("abx_cd", [
    ["equal", 0, 2, 0, 2],
    ["delete", 2, 3, 2, 2],
    ["equal", 3, 5, 2, 4],
  ]));

  it("docstring example qabxcd/abycdf", () => expectOp("qabxcd_abycdf", [
    ["delete", 0, 1, 0, 0],
    ["equal", 1, 3, 0, 2],
    ["replace", 3, 4, 2, 3],
    ["equal", 4, 6, 3, 5],
    ["insert", 6, 6, 5, 6],
  ]));

  it("empty and one-sided empties", () => {
    expectOp("empty_both", []);
    expectOp("a_to_empty", [["delete", 0, 1, 0, 0]]);
    expectOp("empty_to_a", [["insert", 0, 0, 0, 1]]);
  });

  it("full equal collapses to one opcode (adjacent-block fold)", () =>
    expectOp("full_equal", [["equal", 0, 3, 0, 3]]));

  it("dup-heavy token arrays pair like difflib", () => {
    expect(sequenceMatcher(["T", "x", "T", "y", "T"], ["T", "y", "T"])).toEqual([
      ["delete", 0, 2, 0, 0],
      ["equal", 2, 5, 0, 3],
    ]);
  });

  it("swap shows difflib's insert-then-delete shape (not a move)", () => {
    expect(sequenceMatcher(["P1", "P2", "P3"], ["P2", "P1", "P3"])).toEqual([
      ["insert", 0, 0, 0, 1],
      ["equal", 0, 1, 1, 2],
      ["delete", 1, 2, 2, 2],
      ["equal", 2, 3, 2, 3],
    ]);
  });

  it("replace runs split at the longest anchor", () =>
    expectOp("replace_run", [
      ["replace", 0, 2, 0, 2],
      ["equal", 2, 3, 2, 3],
      ["replace", 3, 4, 3, 4],
    ]));
});
