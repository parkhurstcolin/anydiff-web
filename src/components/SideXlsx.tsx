"use client";

import { useEffect, useMemo, useRef, useState, type Ref } from "react";

import type { GridRow, SheetPane, XlsxTarget } from "../core/sideView";

const SIDE = "bg-emerald-100 text-emerald-900";
const LEFT = "bg-red-100 text-red-900";
const CHIP = "rounded bg-amber-100 px-1 py-0.5 text-[10px] font-medium text-amber-900";
const GUTTER = "text-[10px] text-zinc-400";
const FOCUS = "relative z-[5] outline outline-2 outline-sky-500";
// Active row/column: the row-number and column-letter headers ARE the
// selectors — click one and the whole line repaints across both halves.
// Changed cells keep their red/green so a highlight never masks a diff;
// odd rows carry a faint stripe so the grid reads as rows and columns
// even before anything is selected.
const SEL_LINE = "bg-sky-200/80";
const HEAD_ON = "bg-sky-200 text-sky-900";
const STRIPE = "bg-zinc-50";
const SEL_CELL = "relative z-[4] outline outline-2 -outline-offset-2 outline-sky-600";

export const xlsxCellId = (pane: number, row: number, col: number, side: "o" | "n") => `sxt-${pane}-${row}-${col}-${side}`;
export const xlsxPaneId = (pane: number) => `sxt-${pane}-h`;

/** Header-selector state: the painted line in this pane. null = axis not
 *  selected — headers pick one line, a cell click paints both. The ring
 *  shows only on the half that was clicked (one selection, one ring);
 *  the painted row/column is positional, so it spans both halves. */
interface Sel {
  pane: number;
  row: number | null;
  col: number | null;
  side: "o" | "n";
}

const colLetter = (n: number): string => {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

function Grid({
  rows,
  col0,
  side,
  pid,
  focus,
  sel,
  boxRef,
  onScroll,
  onLine,
  onPick,
}: {
  rows: GridRow[];
  col0: number;
  side: "old" | "new";
  pid: number;
  focus: XlsxTarget | null;
  sel: Sel | null;
  boxRef?: Ref<HTMLDivElement>;
  onScroll?: () => void;
  onLine?: (row: number | null, col: number | null, side: "o" | "n") => void;
  onPick?: (t: XlsxTarget) => void;
}) {
  const changed = side === "old" ? LEFT : SIDE;
  const sKey = side === "old" ? ("o" as const) : ("n" as const); // half key as stored in Sel
  const selHere = sel && sel.pane === pid ? sel : null;
  const letters: string[] = [];
  for (let c = col0; c < col0 + (rows[0]?.cells.length ?? 0); c++) letters.push(colLetter(c));
  return (
    // no height cap: the grid grows with the sheet and the page scrolls
    <div ref={boxRef} onScroll={onScroll} className="overflow-x-auto rounded border border-zinc-200">
      <table className="border-collapse">
        <thead>
          <tr className="text-[11px]">
            <th className="sticky left-0 top-0 z-20 border-b border-r border-zinc-300 bg-zinc-100 px-1.5 py-1" />
            {letters.map((l, ci) => (
              <th
                key={l}
                onClick={() => onLine?.(null, ci, sKey)} // the selector: paint this whole column
                title={`highlight column ${l}`}
                className={`sticky top-0 z-10 min-w-16 cursor-pointer border-b border-r px-2 py-1 font-semibold select-none ${
                  selHere?.col === ci
                    ? `border-sky-300 ${HEAD_ON}`
                    : "border-zinc-300 bg-zinc-100 text-zinc-500 hover:bg-zinc-200"
                }`}
              >
                {l}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, ri) => (
            <tr key={r.n}>
              <th
                onClick={() => onLine?.(ri, null, sKey)} // the selector: paint this whole row
                title={`highlight row ${r.n}`}
                className={`sticky left-0 z-10 cursor-pointer border-b border-r px-1.5 py-0.5 text-right align-top text-[10px] font-semibold select-none ${
                  selHere?.row === ri
                    ? `border-sky-300 ${HEAD_ON}`
                    : "border-zinc-300 bg-zinc-100 text-zinc-500 hover:bg-zinc-200"
                }`}
              >
                {r.n}
              </th>
              {r.cells.map((cell, ci) => {
                // rings mark a REAL cell — paint one only on the half where
                // its content exists AND that was clicked; on the other half
                // the painted lines still show its position, but no phantom
                // selector sits stuck on a blank or mirrored cell
                const here = !!cell;
                const isSel = here && selHere?.side === sKey && selHere?.row === ri && selHere?.col === ci;
                // one background per cell, honest precedence: diff color
                // > active line > idle stripe — a highlight never masks a diff
                const bg = cell?.changed
                  ? changed
                  : selHere?.row === ri || selHere?.col === ci
                    ? SEL_LINE
                    : ri % 2 === 1
                      ? STRIPE
                      : "";
                return (
                  <td
                    key={ci}
                    id={xlsxCellId(pid, ri, ci, sKey)}
                    title={cell?.title ?? undefined}
                    onClick={() => {
                      onLine?.(ri, ci, sKey);
                      if (cell?.changed && onPick) onPick({ pane: pid, row: ri, col: ci });
                    }}
                    className={`max-w-56 cursor-pointer truncate border-b border-r border-zinc-200 px-2 py-0.5 font-mono text-xs ${
                      cell?.changed ? changed : "text-zinc-700"
                    } ${bg} ${isSel ? SEL_CELL : ""} ${
                      // while a manual selection is live it owns the ring —
                      // the navigator focus only surfaces on its own jumps
                      here && !selHere && focus && focus.pane === pid && focus.row === ri && focus.col === ci ? FOCUS : ""
                    }`}
                  >
                    {cell?.text ?? ""}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// Vertical is one shared page scroll now; horizontal still needs the two
// halves locked together so a column never drifts out of line.
function syncScrollLeft(src: HTMLDivElement | null, dst: HTMLDivElement | null, lock: { current: boolean }): void {
  if (!src || !dst || lock.current || dst.scrollLeft === src.scrollLeft) return;
  lock.current = true;
  dst.scrollLeft = src.scrollLeft;
  requestAnimationFrame(() => {
    lock.current = false;
  });
}

function Pane({
  pane,
  pid,
  focus,
  sel,
  onLine,
  onPick,
}: {
  pane: SheetPane;
  pid: number;
  focus: XlsxTarget | null;
  sel: Sel | null;
  onLine: (row: number | null, col: number | null, side: "o" | "n") => void;
  onPick: (t: XlsxTarget) => void;
}) {
  const oldBox = useRef<HTMLDivElement>(null);
  const newBox = useRef<HTMLDivElement>(null);
  const syncing = useRef(false);
  const paired = pane.kind === "pair"; // both halves share the column window

  const headTarget: XlsxTarget = { pane: pid, row: -1, col: -1 };
  const isHead = focus && focus.pane === pid && focus.row === -1;
  const label = (s: string) => <span className="truncate font-mono text-xs text-zinc-700">{s || "—"}</span>;
  return (
    <div className="mt-4 overflow-hidden rounded-lg border border-zinc-200 bg-white">
      <div className="flex flex-wrap items-center gap-2 border-b border-zinc-200 bg-zinc-50 px-3 py-1.5">
        {pane.rename ? (
          <span id={xlsxPaneId(pid)} className={`text-xs text-zinc-700 ${isHead ? "rounded outline outline-2 outline-sky-500" : ""}`}>
            <span className={`rounded px-1 ${LEFT}`}>{pane.rename.old}</span>
            <span className="text-zinc-400"> → </span>
            <span className={`rounded px-1 ${SIDE}`}>{pane.rename.new}</span>
          </span>
        ) : (
          <div className="grid flex-1 grid-cols-2 gap-4">
            {label(pane.oldLabel)}
            {label(pane.newLabel)}
          </div>
        )}
        {pane.rename && pane.confidence === "moved" && <span className={CHIP}>renamed, moved</span>}
        {pane.rename && pane.confidence === "manual-inspect" && <span className={CHIP}>renamed, inspect</span>}
        {pane.note && (
          <span id={pane.rename ? undefined : xlsxPaneId(pid)} className={`${CHIP} ${isHead && !pane.rename ? "outline outline-2 outline-sky-500" : ""}`}>
            {pane.note}
          </span>
        )}
        {(pane.rename || pane.note) && (
          <button className="text-[10px] text-sky-600 hover:underline" onClick={() => onPick(headTarget)}>
            jump
          </button>
        )}
        {pane.truncated && <span className={GUTTER}>cut off at 2,000 rows / 26 columns</span>}
      </div>
      <div className="grid grid-cols-2 gap-4 p-3">
        {pane.kind === "added" ? (
          <div className={GUTTER}>not in before</div>
        ) : (
          <Grid rows={pane.oldRows} col0={pane.col0} side="old" pid={pid} focus={focus} sel={sel} boxRef={oldBox} onScroll={paired ? () => syncScrollLeft(oldBox.current, newBox.current, syncing) : undefined} onLine={onLine} onPick={onPick} />
        )}
        {pane.kind === "removed" ? (
          <div className={GUTTER}>not in after</div>
        ) : (
          <Grid rows={pane.newRows} col0={pane.col0} side="new" pid={pid} focus={focus} sel={sel} boxRef={newBox} onScroll={paired ? () => syncScrollLeft(newBox.current, oldBox.current, syncing) : undefined} onLine={onLine} onPick={onPick} />
        )}
      </div>
    </div>
  );
}

export default function SideXlsx({
  panes,
  targets,
  focus,
  onPick,
  tab,
  onTab,
}: {
  panes: SheetPane[];
  targets: XlsxTarget[];
  focus: XlsxTarget | null;
  onPick: (index: number) => void;
  tab: number;
  onTab: (i: number) => void;
}) {
  const indexOf = useMemo(() => {
    const m = new Map<string, number>();
    targets.forEach((t, i) => m.set(`${t.pane}:${t.row}:${t.col}`, i));
    return m;
  }, [targets]);
  const counts = useMemo(() => {
    const c = new Map<number, number>();
    targets.forEach((t) => c.set(t.pane, (c.get(t.pane) ?? 0) + 1));
    return c;
  }, [targets]);

  const [lastPanes, setLastPanes] = useState(panes);
  const [lastFocus, setLastFocus] = useState(focus);
  const [sel, setSel] = useState<Sel | null>(null);
  if (lastPanes !== panes) {
    // fresh diff → the old coordinates may point at unrelated cells
    setLastPanes(panes);
    if (sel) setSel(null);
  }
  if (lastFocus !== focus) {
    // a navigator jump retires a stale manual selection, so the click ring
    // never trails behind the jump; a click on a changed cell sets focus and
    // sel together on the same cell — that pair stays
    setLastFocus(focus);
    if (focus && sel && (sel.row !== focus.row || sel.col !== focus.col)) setSel(null);
  }
  // the selector: a cell click paints its row AND column and rings the
  // half that was clicked; a header click paints just that line; clicking
  // the same line in the same half again clears it
  const pickLine = (row: number | null, col: number | null, side: "o" | "n") =>
    setSel((s) =>
      s && s.pane === tab && s.row === row && s.col === col && s.side === side ? null : { pane: tab, row, col, side },
    );

  const lastJump = useRef<XlsxTarget | null>(null);
  useEffect(() => {
    if (focus === lastJump.current) return; // only move on a fresh jump, not on tab clicks or re-renders
    lastJump.current = focus;
    if (!focus || focus.pane !== tab) return;
    const id = focus.row === -1 ? xlsxPaneId(focus.pane) : xlsxCellId(focus.pane, focus.row, focus.col, "o");
    const alt = focus.row === -1 ? id : xlsxCellId(focus.pane, focus.row, focus.col, "n");
    requestAnimationFrame(() => {
      (document.getElementById(id) ?? document.getElementById(alt))?.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
    });
  }, [focus, tab]);

  const pick = (t: XlsxTarget) => {
    const i = indexOf.get(`${t.pane}:${t.row}:${t.col}`);
    if (i !== undefined) onPick(i);
  };

  const showTabs = panes.length > 1;
  const visible = showTabs ? [panes[tab] ?? panes[0]] : panes;
  return (
    <>
      {showTabs && (
        <div className="mt-3 flex flex-wrap items-end gap-1 border-b border-zinc-200">
          {panes.map((p, i) => {
            const n = counts.get(i) ?? 0;
            const active = i === tab;
            return (
              <button
                key={i}
                onClick={() => onTab(i)}
                className={`flex items-center gap-1.5 rounded-t px-2.5 py-1 text-xs ${
                  active ? "border-x border-t border-zinc-200 bg-white font-medium text-zinc-800" : "text-zinc-500 hover:text-zinc-700"
                }`}
              >
                <span className="max-w-40 truncate font-mono">{p.newLabel || p.oldLabel}</span>
                {n > 0 && <span className={`rounded-full px-1.5 text-[10px] ${active ? "bg-red-100 text-red-900" : "bg-zinc-100 text-zinc-500"}`}>{n}</span>}
              </button>
            );
          })}
        </div>
      )}
      {visible.map((p) => (
        <Pane key={panes.indexOf(p)} pane={p} pid={panes.indexOf(p)} focus={focus} sel={sel} onLine={pickLine} onPick={pick} />
      ))}
    </>
  );
}
