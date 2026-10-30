"use client";

import { useState } from "react";

import SideBySideView from "./SideBySideView";
import { fmtValue, visible } from "../core/format";
import type { Change } from "../core/model";
import type { DocPair } from "../core/sideView";
import { diffWords } from "../core/wordDiff";

const CAP = 40;

export const ADDED = "rounded bg-emerald-100 px-0.5 text-emerald-900";
export const REMOVED = "rounded bg-red-100 px-0.5 text-red-900 line-through decoration-red-400";
const AMBER = "rounded bg-amber-100 px-1 py-0.5 text-amber-900";
const NEUTRAL = "rounded bg-zinc-100 px-1 text-zinc-600";

function ValueDiff({ c }: { c: Change }) {
  if (typeof c.before === "string" && typeof c.after === "string") {
    return (
      <>
        {diffWords(c.before, c.after).map((s, i) =>
          s.kind === "same" ? (
            <span key={i}>{visible(s.text)}</span>
          ) : (
            <span key={i} className={s.kind === "added" ? ADDED : REMOVED}>
              {visible(s.text)}
            </span>
          ),
        )}
      </>
    );
  }
  const side = (v: string | number | boolean | null) =>
    v === null ? "" : typeof v === "string" ? `"${visible(fmtValue(v))}"` : fmtValue(v);
  return (
    <>
      {c.before !== null && <span className={REMOVED}>{side(c.before)}</span>}
      {c.before !== null && c.after !== null && <span className="text-zinc-400"> → </span>}
      {c.after !== null && <span className={ADDED}>{side(c.after)}</span>}
    </>
  );
}

function rowLabel(c: Change): { badge: string; style: string } {
  if (c.kind === "sheet_added") return { badge: "sheet added", style: ADDED };
  if (c.kind === "sheet_removed") return { badge: "sheet removed", style: REMOVED };
  if (c.kind === "sheet_renamed") return { badge: "renamed", style: AMBER };
  if (c.before === null) return { badge: "added", style: ADDED };
  if (c.after === null) return { badge: "removed", style: REMOVED };
  if (c.before === c.after) {
    // text held; something else moved: paragraph style or run formatting
    return { badge: c.note?.includes("style") ? "style" : "format", style: AMBER };
  }
  return { badge: "edit", style: NEUTRAL };
}

const isStructural = (c: Change) => c.kind.startsWith("sheet_");

function groupKey(c: Change): string {
  if (c.kind === "cell") return c.where.includes("!") ? c.where.slice(0, c.where.indexOf("!")) : "—";
  const t = /^table (\d+)/.exec(c.where);
  if (t) return `Table ${t[1]}`;
  return "Document";
}

export default function ResultsView({
  changes,
  pair,
  mode,
  onMode,
  oldName,
  newName,
}: {
  changes: Change[];
  pair?: DocPair | null;
  mode: "side" | "list";
  onMode: (m: "side" | "list") => void;
  oldName?: string;
  newName?: string;
}) {
  const [showAll, setShowAll] = useState(false);
  const structural = changes.filter(isStructural);
  const content = changes.filter((c) => !isStructural(c));
  const groups = new Map<string, Change[]>();
  for (const c of content) {
    const k = groupKey(c);
    const arr = groups.get(k);
    if (arr) arr.push(c);
    else groups.set(k, [c]);
  }
  const edits = content.filter((c) => c.before !== null && c.after !== null).length;
  const adds = content.filter((c) => c.before === null).length;
  const removes = content.filter((c) => c.after === null).length;
  const seg = (active: boolean) =>
    `px-3 py-1.5 ${active ? "bg-zinc-900 text-white" : "bg-white text-zinc-600 hover:bg-zinc-100"}`;
  const sideActive = mode === "side" && !!pair && changes.length > 0;

  return (
    <section className={sideActive ? "mt-3" : "mt-8"}>
      <div className="flex flex-wrap items-baseline gap-x-4 gap-y-1">
        <h2 className="text-sm font-semibold text-zinc-900">
          {changes.length === 0
            ? "Semantically identical"
            : `${edits} edit${edits === 1 ? "" : "s"} · ${adds} added · ${removes} removed${
                structural.length ? ` · ${structural.length} sheet-level` : ""
              }`}
        </h2>
        <span className="text-xs text-zinc-500">
          {oldName} → {newName}
        </span>
        {changes.length > 0 && pair && (
          <div className="ml-auto flex overflow-hidden rounded-md border border-zinc-300 text-xs font-medium">
            <button type="button" className={seg(mode === "side")} onClick={() => onMode("side")}>
              Side by side
            </button>
            <button type="button" className={`border-l border-zinc-300 ${seg(mode === "list")}`} onClick={() => onMode("list")}>
              Change list
            </button>
          </div>
        )}
      </div>

      {changes.length === 0 && (
        <p className="mt-3 rounded-md bg-emerald-50 p-3 text-sm text-emerald-800">
          Every sheet, value, formula, paragraph and table cell matches. Formatting-only edits are
          deliberately invisible.
        </p>
      )}

      {mode === "side" && pair && changes.length > 0 && (
        <SideBySideView pair={pair} changes={changes} oldName={oldName} newName={newName} />
      )}

      {!sideActive && (
        <>
          {structural.length > 0 && (
        <div className="mt-4 rounded-lg border border-zinc-200 bg-white">
          {structural.map((c, i) => {
            const l = rowLabel(c);
            return (
              <div key={i} className="flex items-center gap-3 border-b border-zinc-100 px-4 py-2 text-sm last:border-b-0">
                <span className={`text-xs font-medium ${l.style}`}>{l.badge}</span>
                <span className="font-mono text-xs">{c.where}</span>
                {c.confidence !== "exact" && <span className={AMBER}>{c.confidence === "moved" ? "moved" : "moved, inspect manually"}</span>}
                {c.kind === "sheet_renamed" && (
                  <span className="text-xs text-zinc-500">
                    {c.before} → {c.after}
                  </span>
                )}
              </div>
            );
          })}
        </div>
      )}

      {[...groups.entries()].map(([group, rows]) => {
        const shown = showAll ? rows : rows.slice(0, CAP);
        return (
          <div key={group} className="mt-4 overflow-hidden rounded-lg border border-zinc-200 bg-white">
            <div className="border-b border-zinc-200 bg-zinc-50 px-4 py-2 text-xs font-semibold text-zinc-700">
              {group} <span className="font-normal text-zinc-500">· {rows.length} change{rows.length === 1 ? "" : "s"}</span>
            </div>
            <table className="w-full border-collapse text-sm">
              <tbody>
                {shown.map((c, i) => {
                  const l = rowLabel(c);
                  return (
                    <tr key={i} className="border-b border-zinc-100 align-top last:border-b-0">
                      <td className="w-24 px-4 py-2">
                        <span className={`text-xs font-medium ${l.style}`}>{l.badge}</span>
                      </td>
                      <td className="w-36 py-2 pr-3 font-mono text-xs text-zinc-700">
                        {c.where.includes("!") ? c.where.slice(c.where.indexOf("!") + 1) : c.where}
                        {c.confidence !== "exact" && (
                          <div className={`mt-1 text-[10px] ${AMBER}`}>
                            {c.confidence === "moved" ? "moved" : "inspect manually"}
                          </div>
                        )}
                      </td>
                      <td className="py-2 pr-4 font-mono text-xs">
                        <ValueDiff c={c} />
                        {c.note && <div className="mt-1 text-[11px] text-zinc-500">{c.note}</div>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {!showAll && rows.length > CAP && (
              <button
                type="button"
                onClick={() => setShowAll(true)}
                className="w-full border-t border-zinc-100 px-4 py-2 text-left text-xs text-sky-700 hover:bg-zinc-50"
              >
                + {rows.length - CAP} more — show all
              </button>
            )}
          </div>
        );
      })}
        </>
      )}
    </section>
  );
}
