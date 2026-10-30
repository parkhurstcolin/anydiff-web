/**
 * The change record — the single contract every renderer consumes,
 * ported from the engine's model.py (BRIEF §3). One contract, two engines:
 * golden tests compare TS output against the Python engine field by field.
 */

export const CHANGE_KINDS = [
  "cell",
  "sheet_added",
  "sheet_removed",
  "sheet_renamed",
  "row",
  "col",
  "paragraph",
  "slide",
  "image",
  "meta",
] as const;

export const CONFIDENCES = ["exact", "moved", "manual-inspect"] as const;

// Types derived from the runtime arrays (typeof X[number]): the validator and
// the type can never drift apart, which the Python __post_init__ cannot promise.
export type ChangeKind = (typeof CHANGE_KINDS)[number];
export type Confidence = (typeof CONFIDENCES)[number];

/** Cell/paragraph payloads are scalars; table rows are joined strings. */
export type ChangeValue = string | number | boolean | null;

export interface Change {
  path: string;
  kind: ChangeKind;
  where: string;
  before: ChangeValue;
  after: ChangeValue;
  confidence: Confidence;
  /**
   * Annotation decided by the differ, which alone knows whether a claim
   * ("value unchanged") was verified against cached values. Renderers must
   * display it verbatim, never infer it from the values.
   */
  note: string | null;
}

/** Mirrors the dataclass defaults: confidence "exact", note null. */
export function change(
  c: Pick<Change, "path" | "kind" | "where" | "before" | "after"> &
    Partial<Pick<Change, "confidence" | "note">>,
): Change {
  const full: Change = { confidence: "exact", note: null, ...c };
  if (!isChange(full)) {
    throw new Error(`invalid change record: ${JSON.stringify(full)}`);
  }
  return full;
}

/** Runtime guard for records crossing an untyped boundary (JSON, oracle files). */
export function isChange(value: unknown): value is Change {
  if (typeof value !== "object" || value === null) return false;
  const v = value as Record<string, unknown>;
  return (
    typeof v.path === "string" &&
    typeof v.kind === "string" &&
    CHANGE_KINDS.includes(v.kind as ChangeKind) &&
    typeof v.where === "string" &&
    (v.before === null || ["string", "number", "boolean"].includes(typeof v.before)) &&
    (v.after === null || ["string", "number", "boolean"].includes(typeof v.after)) &&
    typeof v.confidence === "string" &&
    CONFIDENCES.includes(v.confidence as Confidence) &&
    (v.note === null || typeof v.note === "string")
  );
}
