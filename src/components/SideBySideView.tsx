"use client";

/** Git-style side-by-side: both documents rendered in place with changes
 *  highlighted in context. Model logic lives in core/sideView.ts; this picks
 *  the format, owns the change navigator (targets, focus, n/p keys). */

import { useEffect, useMemo, useState } from "react";

import { docxLineId } from "./SideDocx";
import SideDocx from "./SideDocx";
import SideXlsx from "./SideXlsx";
import type { Change } from "../core/model";
import { docxTargets, sideDocx, sideXlsx, xlsxTargets, type DocPair, type XlsxTarget } from "../core/sideView";

const MAX_DOCX_LINES = 400;

export default function SideBySideView({
  pair,
  changes,
  oldName,
  newName,
  computed,
}: {
  pair: DocPair;
  changes: Change[];
  oldName?: string;
  newName?: string;
  computed?: boolean; // values came from the browser recalculation
}) {
  const [nav, setNav] = useState({ pair, focus: 0, tab: 0, touched: false });
  if (nav.pair !== pair) setNav({ pair, focus: 0, tab: 0, touched: false }); // fresh diff: start at the top

  const lines = useMemo(() => (pair.kind === "docx" ? sideDocx(pair.old, pair.new) : null), [pair]);
  const shown = useMemo(() => lines?.slice(0, MAX_DOCX_LINES) ?? null, [lines]);
  const panes = useMemo(() => (pair.kind === "xlsx" ? sideXlsx(changes, pair.old, pair.new, { values: computed }) : []), [pair, changes, computed]);
  const targets = useMemo(
    () => (pair.kind === "xlsx" ? xlsxTargets(panes) : docxTargets(shown ?? [])),
    [pair.kind, panes, shown],
  );

  const count = targets.length;
  const at = count ? ((nav.focus % count) + count) % count : 0;
  const go = (i: number) =>
    setNav((s) => {
      const n = ((i % count) + count) % count;
      const tab = pair.kind === "xlsx" ? (targets[n] as XlsxTarget | undefined)?.pane ?? s.tab : s.tab;
      return { ...s, focus: n, tab, touched: true };
    });

  useEffect(() => {
    if (count < 2) return;
    const step = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement | null;
      if (e.metaKey || e.ctrlKey || e.altKey || (t && (["INPUT", "TEXTAREA", "SELECT"].includes(t.tagName) || t.isContentEditable))) return;
      if (e.key === "n" || e.key === "ArrowRight" || e.key === "j") {
        e.preventDefault();
        go(at + 1);
      } else if (e.key === "p" || e.key === "ArrowLeft" || e.key === "k") {
        e.preventDefault();
        go(at - 1);
      }
    };
    window.addEventListener("keydown", step);
    return () => window.removeEventListener("keydown", step);
  });

  useEffect(() => {
    if (!nav.touched || pair.kind !== "docx" || !count) return;
    document.getElementById(docxLineId(targets[at] as number))?.scrollIntoView({ block: "center", behavior: "smooth" });
  }, [nav.touched, pair.kind, targets, at, count]);

  const xfocus = pair.kind === "xlsx" && nav.touched && count ? (targets[at] as XlsxTarget) : null;
  const dfocus = pair.kind === "docx" && nav.touched && count ? (targets[at] as number) : null;

  return (
    <div className="mt-3">
      <div className="sticky top-0 z-30 -mx-4 mb-1 flex flex-wrap items-center justify-between gap-2 border-b border-zinc-200 bg-white/95 px-4 py-1.5 backdrop-blur">
        <div className="flex flex-wrap gap-3 text-[11px] text-zinc-500">
          <span className="rounded bg-red-100 px-2 py-0.5 text-red-900">before</span>
          <span className="rounded bg-emerald-100 px-2 py-0.5 text-emerald-900">after</span>
          <span className="hidden sm:inline">context is dimmed · click a highlight · n / p to step · click a row number / column letter to trace its line</span>
        </div>
        {count > 0 && (
          <div className="flex items-center gap-1.5 text-xs">
            <button
              onClick={() => go(at - 1)}
              className="rounded border border-zinc-200 px-2 py-0.5 text-zinc-600 hover:bg-zinc-50"
              title="previous change (p)"
            >
              ‹
            </button>
            <span className="min-w-16 text-center tabular-nums text-zinc-600">
              {at + 1} / {count}
            </span>
            <button
              onClick={() => go(at + 1)}
              className="rounded border border-zinc-200 px-2 py-0.5 text-zinc-600 hover:bg-zinc-50"
              title="next change (n)"
            >
              ›
            </button>
          </div>
        )}
      </div>

      <div className="grid grid-cols-2 gap-4 text-[11px] font-medium text-zinc-500">
        <span className="truncate">before · {oldName}</span>
        <span className="truncate">after · {newName}</span>
      </div>

      {pair.kind === "xlsx" ? (
        <SideXlsx panes={panes} targets={targets as XlsxTarget[]} focus={xfocus} onPick={go} tab={nav.tab} onTab={(i) => setNav((s) => ({ ...s, tab: i }))} />
      ) : (
        <div className="mt-3">
          {shown && <SideDocx lines={shown} targets={targets as number[]} focusLine={dfocus} onPick={go} />}
          {lines && shown && lines.length > shown.length && (
            <div className="mt-1 rounded-md bg-zinc-50 px-3 py-2 text-xs text-zinc-500">
              + {lines.length - shown.length} more lines — see the change list
            </div>
          )}
        </div>
      )}
    </div>
  );
}
