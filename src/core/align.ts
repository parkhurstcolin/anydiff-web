/**
 * Hash-anchored alignment — port of the engine's align.py (BRIEF §3 rule 1).
 * A renamed-or-edited sheet, or a moved slide, must pair with its old self
 * instead of showing removed+added. Greedy Jaccard over content token sets;
 * weak or contested matches degrade to "manual-inspect", every re-paired
 * item is at best "moved" — never a silent false positive.
 */

export const MIN_SIM = 0.5;
export const STRONG_SIM = 0.8;
export const RIVAL_MARGIN = 0.05;

export type TokenKey = string | number;
/** Eager token set, or a zero-arg thunk computed only if content anchoring
 *  is reached (exact-name matches must not pay to tokenize big sheets). */
export type TokenSource = ReadonlySet<string> | (() => ReadonlySet<string>);

export interface Pair<K> {
  old: K;
  new: K;
  confidence: "exact" | "moved" | "manual-inspect";
}

export interface Alignment<K> {
  pairs: Pair<K>[];
  removed: K[];
  added: K[];
}

/** Python-order comparator: numbers numerically, strings by code point
 *  (JS default sort is stringly — "9" > "10"; localeCompare is locale-aware
 *  and would disagree with Python's str ordering). */
export function keyCmp(a: TokenKey, b: TokenKey): number {
  if (typeof a === "number" && typeof b === "number") return a - b;
  const [as, bs] = [String(a), String(b)];
  return as < bs ? -1 : as > bs ? 1 : 0;
}

export function jaccard(a: ReadonlySet<string>, b: ReadonlySet<string>): number {
  if (a.size === 0 && b.size === 0) return 1.0;
  let inter = 0;
  for (const t of a) if (b.has(t)) inter++;
  const union = a.size + b.size - inter;
  return union ? inter / union : 1.0;
}

export function alignByAnchors<K extends TokenKey>(
  oldTokens: Map<K, TokenSource>,
  newTokens: Map<K, TokenSource>,
  anchorExactKeys = true,
): Alignment<K> {
  const pairs: Pair<K>[] = [];
  let oldLeft: K[];
  let newLeft: K[];
  if (anchorExactKeys) {
    const shared = new Set<K>();
    for (const k of oldTokens.keys()) if (newTokens.has(k)) shared.add(k);
    for (const key of [...shared].sort(keyCmp)) {
      pairs.push({ old: key, new: key, confidence: "exact" });
    }
    oldLeft = [...oldTokens.keys()].filter((k) => !shared.has(k)).sort(keyCmp);
    newLeft = [...newTokens.keys()].filter((k) => !shared.has(k)).sort(keyCmp);
  } else {
    oldLeft = [...oldTokens.keys()].sort(keyCmp);
    newLeft = [...newTokens.keys()].sort(keyCmp);
  }

  // resolve each leftover's tokens ONCE (a thunk would re-tokenize per candidate pair)
  const resolve = (v: TokenSource): ReadonlySet<string> =>
    typeof v === "function" ? v() : v;
  const oldRes = new Map<K, ReadonlySet<string>>(
    oldLeft.map((o) => [o, resolve(oldTokens.get(o)!)]),
  );
  const newRes = new Map<K, ReadonlySet<string>>(
    newLeft.map((n) => [n, resolve(newTokens.get(n)!)]),
  );

  const sims = new Map<string, { o: K; n: K; sim: number }>();
  for (const o of oldLeft) {
    for (const n of newLeft) {
      sims.set(`${o}\u0000${n}`, { o, n, sim: jaccard(oldRes.get(o)!, newRes.get(n)!) });
    }
  }

  const candidates = [...sims.values()].sort(
    (x, y) => y.sim - x.sim || keyCmp(x.o, y.o) || keyCmp(x.n, y.n),
  );
  const takenOld = new Set<K>();
  const takenNew = new Set<K>();
  const matched: { o: K; n: K; sim: number }[] = [];
  for (const c of candidates) {
    if (c.sim < MIN_SIM || takenOld.has(c.o) || takenNew.has(c.n)) continue;
    takenOld.add(c.o);
    takenNew.add(c.n);
    matched.push(c);
  }

  for (const { o, n, sim } of matched.sort(
    (x, y) => keyCmp(x.o, y.o) || keyCmp(x.n, y.n),
  )) {
    let contested = false;
    for (const other of sims.values()) {
      if (other === sims.get(`${o}\u0000${n}`)) continue;
      if ((other.o === o || other.n === n) && other.sim >= sim - RIVAL_MARGIN) {
        contested = true;
        break;
      }
    }
    pairs.push({
      old: o,
      new: n,
      confidence: contested || sim < STRONG_SIM ? "manual-inspect" : "moved",
    });
  }

  return {
    pairs,
    removed: oldLeft.filter((k) => !takenOld.has(k)),
    added: newLeft.filter((k) => !takenNew.has(k)),
  };
}
