/**
 * Word-level before/after highlighting for the UI (display-only: Change
 * records stay engine-identical; this never feeds back into the diff).
 * Token = run of non-whitespace or run of whitespace, so text rebuilds exactly.
 */

export type SegKind = "same" | "added" | "removed";
export interface Seg {
  text: string;
  kind: SegKind;
}

const MAX_SIDE = 400; // LCS is O(n*m); long cells fall back to whole-side marks

export function tokenize(s: string): string[] {
  return s.split(/(\s+)/).filter((t) => t.length > 0);
}

export function diffWords(before: string, after: string): Seg[] {
  const a = tokenize(before);
  const b = tokenize(after);
  const whole: Seg[] = [];
  if (a.length) whole.push({ text: before, kind: "removed" });
  if (b.length) whole.push({ text: after, kind: "added" });
  if (!a.length || !b.length) return whole;
  if (a.length > MAX_SIDE || b.length > MAX_SIDE) return whole;

  const n = a.length;
  const m = b.length;
  const lcs = new Uint32Array((n + 1) * (m + 1));
  const at = (i: number, j: number) => i * (m + 1) + j;
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[at(i, j)] =
        a[i] === b[j] ? lcs[at(i + 1, j + 1)] + 1 : Math.max(lcs[at(i + 1, j)], lcs[at(i, j + 1)]);
    }
  }
  const segs: Seg[] = [];
  const push = (text: string, kind: SegKind) => {
    const last = segs[segs.length - 1];
    if (last && last.kind === kind) last.text += text;
    else segs.push({ text, kind });
  };
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      push(a[i], "same");
      i++;
      j++;
    } else if (lcs[at(i + 1, j)] >= lcs[at(i, j + 1)]) {
      push(a[i++], "removed");
    } else {
      push(b[j++], "added");
    }
  }
  while (i < n) push(a[i++], "removed");
  while (j < m) push(b[j++], "added");
  return segs;
}
