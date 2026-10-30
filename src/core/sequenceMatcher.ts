/**
 * Faithful TS port of difflib.SequenceMatcher with autojunk=False and no
 * junk function (the engine's exact configuration). The DOCX differ pairs
 * paragraphs and table rows by its opcodes, so results must match Python
 * element-for-element — tests pin against difflib-generated expectations.
 */

export type OpcodeTag = "replace" | "delete" | "insert" | "equal";
export type Opcode = readonly [tag: OpcodeTag, i1: number, i2: number, j1: number, j2: number];

export function sequenceMatcher<T>(
  a: readonly T[],
  b: readonly T[],
): Opcode[] {
  const b2j = new Map<T, number[]>();
  b.forEach((el, j) => {
    const arr = b2j.get(el);
    if (arr) arr.push(j);
    else b2j.set(el, [j]);
  });

  // longest match in a[alo:ahi] x b[blo:bhi]; no junk (autojunk=False), so
  // the "extend by junk" passes of the Python original are provably no-ops
  const findLongestMatch = (alo: number, ahi: number, blo: number, bhi: number) => {
    let besti = alo;
    let bestj = blo;
    let bestsize = 0;
    let j2len = new Map<number, number>();
    for (let i = alo; i < ahi; i++) {
      const newj2len = new Map<number, number>();
      const js = b2j.get(a[i]);
      if (js) {
        for (const j of js) {
          if (j < blo) continue;
          if (j >= bhi) break;
          const k = (j2len.get(j - 1) ?? 0) + 1;
          newj2len.set(j, k);
          if (k > bestsize) {
            besti = i - k + 1;
            bestj = j - k + 1;
            bestsize = k;
          }
        }
      }
      j2len = newj2len;
    }
    while (besti > alo && bestj > blo && a[besti - 1] === b[bestj - 1]) {
      besti--;
      bestj--;
      bestsize++;
    }
    while (
      besti + bestsize < ahi &&
      bestj + bestsize < bhi &&
      a[besti + bestsize] === b[bestj + bestsize]
    ) {
      bestsize++;
    }
    return [besti, bestj, bestsize] as const;
  };

  const la = a.length;
  const lb = b.length;
  const queue: [number, number, number, number][] = [[0, la, 0, lb]];
  const blocks: [number, number, number][] = [];
  while (queue.length) {
    const [alo, ahi, blo, bhi] = queue.pop()!;
    const [i, j, k] = findLongestMatch(alo, ahi, blo, bhi);
    if (k) {
      blocks.push([i, j, k]);
      if (alo < i && blo < j) queue.push([alo, i, blo, j]);
      if (i + k < ahi && j + k < bhi) queue.push([i + k, ahi, j + k, bhi]);
    }
  }
  blocks.sort((p, q) => p[0] - q[0] || p[1] - q[1] || p[2] - q[2]);

  // collapse adjacent equal blocks (difflib 2.5+ behavior)
  const nonAdjacent: [number, number, number][] = [];
  let i1 = 0;
  let j1 = 0;
  let k1 = 0;
  for (const [i2, j2, k2] of blocks) {
    if (i1 + k1 === i2 && j1 + k1 === j2) {
      k1 += k2;
    } else {
      if (k1) nonAdjacent.push([i1, j1, k1]);
      i1 = i2;
      j1 = j2;
      k1 = k2;
    }
  }
  if (k1) nonAdjacent.push([i1, j1, k1]);
  nonAdjacent.push([la, lb, 0]); // sentinel: makes the tail emit its gap opcode

  const answer: Opcode[] = [];
  let i = 0;
  let j = 0;
  for (const [ai, bj, size] of nonAdjacent) {
    let tag: OpcodeTag | "" = "";
    if (i < ai && j < bj) tag = "replace";
    else if (i < ai) tag = "delete";
    else if (j < bj) tag = "insert";
    if (tag) answer.push([tag, i, ai, j, bj]);
    i = ai + size;
    j = bj + size;
    if (size) answer.push(["equal", ai, i, bj, j]);
  }
  return answer;
}
