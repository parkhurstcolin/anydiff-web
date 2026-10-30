/**
 * Human-scale value formatting — port of the engine's format.py.
 * Deterministic rule (demo line `4.2M -> 4.8M` demands it): |v| >= 1e4 gets
 * a K/M/B suffix at 1 decimal with trailing ".0" stripped; everything below
 * stays raw. Consequence: 999_999 renders "1000K", not "1M" — honest
 * rounding beats scale promotion; boundaries pinned by tests.
 *
 * Date lives OUTSIDE the Change contract: extractors call fmtValue on cell
 * values before building records, so Changes stay JSON-clean. (Known
 * boundary vs Python: float-to-string of exotic values like 1e-7 may differ
 * in exponent notation; office fixtures carry no such values.)
 */

// Cc controls (ESC/BEL), Cf format chars (RLO/ZWSP), Cs surrogates, Zl/Zp
// line separators — inert inside a file, actions on a screen. Combining
// marks (Mn) stay visible: they are legitimate rendered text.
const INVISIBLE = /\p{Cc}|\p{Cf}|\p{Cs}|\p{Zl}|\p{Zp}/u;

const SIMPLE_ESCAPES: Record<string, string> = { "\n": "\\n", "\t": "\\t", "\r": "\\r" };

/** Single-line printable view of document-sourced text; invisible chars
 *  become visible escape text (`\x1b`, `\u202e`) — stripping would hide a
 *  change, which is the worse failure. Idempotent: renderers may compose it. */
export function visible(text: string): string {
  if (!INVISIBLE.test(text)) return text;
  return [...text]
    .map((ch) => {
      if (!INVISIBLE.test(ch)) return ch;
      const simple = SIMPLE_ESCAPES[ch];
      if (simple) return simple;
      const cp = ch.codePointAt(0)!;
      return cp < 0x100 ? `\\x${cp.toString(16).padStart(2, "0")}` : `\\u${cp.toString(16).padStart(4, "0")}`;
    })
    .join("");
}

function scaled(v: number, scale: number, unit: string): string {
  const text = (v / scale).toFixed(1);
  return (text.endsWith(".0") ? text.slice(0, -2) : text) + unit;
}

export function fmtValue(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") {
    const av = Math.abs(v);
    if (av >= 1e9) return scaled(v, 1e9, "B");
    if (av >= 1e6) return scaled(v, 1e6, "M");
    if (av >= 1e4) return scaled(v, 1e3, "K");
    return String(v);
  }
  if (v instanceof Date) {
    const p = (n: number) => String(n).padStart(2, "0");
    const d = `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
    const hms = `${p(v.getHours())}:${p(v.getMinutes())}:${p(v.getSeconds())}`;
    return hms === "00:00:00" ? d : `${d} ${hms}`;
  }
  return String(v);
}
