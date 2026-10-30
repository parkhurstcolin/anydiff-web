/**
 * Content fingerprints for instant-equality fast paths (git works this way:
 * compare hashes, descend only when they differ).
 *
 * The rule that keeps this honest: an UNEQUAL hash never decides anything —
 * the full compare runs anyway. Only an EQUAL hash lets work be skipped,
 * and the skip is display-safe at worst: for diff results the fingerprint
 * covers every field that the differ reads (address, value, formula /
 * style, run text, run format, cell text), so a match means "no change can
 * exist" short of a ~2^-64 birthday collision across a whole document.
 * Folding is commutative (sum + xor + count), so order can never hide a
 * difference; ordering INSIDE a block is baked into each item's hash.
 */

export function cyrb53(str: string, seed = 0): number {
  let h1 = 0xdeadbeef ^ seed;
  let h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
  h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
  h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return 4294967296 * (2097151 & h2) + (h1 >>> 0);
}

/** Raw bytes: chunk-hash chained so order and length both matter. */
export function hashBytes(bytes: Uint8Array): number {
  let h = 0;
  const CHUNK = 8192;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    const part = String.fromCharCode(...bytes.subarray(i, i + CHUNK));
    h = cyrb53(part, h);
  }
  return cyrb53(`len:${bytes.length}`, h);
}

/** Commutative accumulator: sum (mod 2^53-1) + xor + count.
 *  `unstable` counts non-self-equal values (NaN): the differ flags those as
 *  changes even against themselves, so any pair with unstable > 0 must run
 *  the full compare — an equal fingerprint may not authorize a skip. */
export interface SumXor {
  sum: number;
  xor: number;
  count: number;
  unstable: number;
}

export const newSumXor = (): SumXor => ({ sum: 0, xor: 0, count: 0, unstable: 0 });

export const sameSumXor = (a: SumXor, b: SumXor): boolean =>
  a.sum === b.sum && a.xor === b.xor && a.count === b.count && a.unstable === 0 && b.unstable === 0;

const fold = (acc: SumXor, h: number): void => {
  acc.sum = (acc.sum + h) % 9007199254740991;
  acc.xor ^= h;
  acc.count += 1;
};

/** One sheet's cells: address + value + formula, order-independent. */
export function foldCell(
  acc: SumXor,
  addr: string,
  value: unknown,
  formula: string | null,
): void {
  let h: number;
  if (value === null || value === undefined) h = 1;
  else if (typeof value === "number") {
    if (Number.isNaN(value)) {
      acc.unstable += 1;
      h = 0xdeadb; // any value: a NaN pair can never authorize a skip
    } else h = cyrb53(`n${value}`, 2);
  } else if (typeof value === "string") h = cyrb53(`s${value}`, 3);
  else if (typeof value === "boolean") h = value ? 104729 : 65537;
  else if (value instanceof Date) h = cyrb53(`d${value.getTime()}`, 5);
  else h = cyrb53(`?${String(value)}`, 8);
  fold(acc, cyrb53(addr, h) ^ (formula === null ? 0x5bf03635 : cyrb53(formula, 9)));
}

/** One docx block: style / run text + bold + italic / table cell text,
 *  each position-salted so swaps and reorderings change the digest. */
export function foldBlock(
  acc: SumXor,
  b: { blockType: string; style?: string; rows?: string[][]; runs?: { text: string; bold: boolean; italic: boolean }[] },
): void {
  if (b.blockType === "table") {
    b.rows?.forEach((row, r) =>
      row.forEach((cell, c) => fold(acc, cyrb53(cell, cyrb53(`t${r}:${c}`, 11)))),
    );
    fold(acc, cyrb53(`rows:${b.rows?.length ?? 0}`, 13));
    return;
  }
  fold(acc, cyrb53(`st:${b.style ?? "Normal"}`, 17));
  b.runs?.forEach((run, i) =>
    fold(acc, cyrb53(`${run.bold ? "b" : ""}${run.italic ? "i" : ""}|${run.text}`, cyrb53(`r${i}`, 19))),
  );
  fold(acc, cyrb53(`runs:${b.runs?.length ?? 0}`, 23));
}
