"use client";

import { diffWords, type Seg } from "../core/wordDiff";
import type { DocLine, TableRowPair } from "../core/sideView";
import type { Run } from "../extract/docx";

const SIDE = "bg-emerald-100 text-emerald-900";
const LEFT = "bg-red-100 text-red-900";
const CHIP = "rounded bg-amber-100 px-1 py-0.5 text-[10px] font-medium text-amber-900";
const STYLE_CHIP = "rounded bg-zinc-100 px-1 text-[10px] text-zinc-500";
const GUTTER = "text-[10px] text-zinc-400";

function Runs({ runs, faint }: { runs: Run[]; faint?: boolean }) {
  return (
    <>
      {runs.map((r, i) => (
        <span key={i} className={`${r.bold ? "font-semibold" : ""} ${r.italic ? "italic" : ""} ${faint ? "text-zinc-500" : "text-zinc-700"}`}>
          {r.text}
        </span>
      ))}
    </>
  );
}

function WordDiff({ segs, mode, fallback }: { segs: Seg[]; mode: "left" | "right"; fallback: string }) {
  const mark = mode === "left" ? LEFT : SIDE;
  if (!segs.length) return <span className={`rounded px-0.5 ${mark}`}>{fallback}</span>;
  return (
    <>
      {segs.map((s, i) => (
        <span key={i} className={s.kind === "same" ? "" : `rounded px-0.5 ${mark}`}>
          {s.text}
        </span>
      ))}
    </>
  );
}

function TRow({ cells, status, side }: { cells: string[] | null; status: TableRowPair["status"]; side: "old" | "new" }) {
  const tone = status === "same" || !cells ? "text-zinc-500" : side === "old" ? LEFT : SIDE;
  return (
    <div className={`min-h-6 rounded px-2 py-0.5 font-mono text-xs ${tone}`}>
      {cells?.map((c, i) => (
        <span key={i}>
          {i > 0 && <span className="text-zinc-300"> | </span>}
          {c || " "}
        </span>
      ))}
    </div>
  );
}

function Line({ line }: { line: DocLine }) {
  const cell = "px-3 py-1.5 text-sm leading-5 align-top";
  if (line.type === "tablePair" || line.type === "tableDump") {
    const rows: TableRowPair[] =
      line.type === "tablePair"
        ? line.rows
        : line.table.rows.map((r) => ({
            status: line.side === "old" ? ("removed" as const) : ("added" as const),
            oldCells: line.side === "old" ? r : null,
            newCells: line.side === "new" ? r : null,
          }));
    return (
      <div className="grid grid-cols-2 gap-4 border-y border-zinc-100 bg-zinc-50/60 py-1">
        <div className="space-y-px">{rows.map((r, i) => <TRow key={i} cells={r.oldCells} status={r.status} side="old" />)}</div>
        <div className="space-y-px">{rows.map((r, i) => <TRow key={i} cells={r.newCells} status={r.status} side="new" />)}</div>
      </div>
    );
  }
  if (line.type === "same") {
    const one = (
      <div className={cell}>
        {line.style !== "Normal" && <span className={`${STYLE_CHIP} mr-1`}>{line.style}</span>}
        <Runs runs={line.runs} faint />
      </div>
    );
    return (
      <div className="grid grid-cols-2 gap-x-4">
        {one}
        {one}
      </div>
    );
  }
  if (line.type === "edit") {
    const chip = line.styleChanged ? (
      <span className={`${CHIP} mr-1`}>
        {line.oldStyle} → {line.newStyle}
      </span>
    ) : line.fmtChanged ? (
      <span className={`${CHIP} mr-1`}>format</span>
    ) : null;
    const segs = diffWords(line.oldText, line.newText);
    return (
      <div className="grid grid-cols-2 gap-x-4 border-y border-zinc-100">
        <div className={`${cell} bg-red-50`}>
          {chip}
          <WordDiff segs={segs.filter((s) => s.kind !== "added")} mode="left" fallback={line.oldText} />
        </div>
        <div className={`${cell} bg-emerald-50`}>
          {chip}
          {line.newStyle !== "Normal" && <span className={`${STYLE_CHIP} mr-1`}>{line.newStyle}</span>}
          <WordDiff segs={segs.filter((s) => s.kind !== "removed")} mode="right" fallback={line.newText} />
        </div>
      </div>
    );
  }
  const removed = line.type === "removed";
  return (
    <div className="grid grid-cols-2 gap-x-4 border-y border-zinc-100">
      <div className={`${cell} ${removed ? "bg-red-50" : ""}`}>
        {removed ? (
          <>
            {line.style !== "Normal" && <span className={`${STYLE_CHIP} mr-1`}>{line.style}</span>}
            <span className="bg-red-100 text-red-900">
              <Runs runs={line.runs} />
            </span>
          </>
        ) : (
          <span className={GUTTER}>not in before</span>
        )}
      </div>
      <div className={`${cell} ${removed ? "" : "bg-emerald-50"}`}>
        {removed ? (
          <span className={GUTTER}>not in after</span>
        ) : (
          <>
            {line.style !== "Normal" && <span className={`${STYLE_CHIP} mr-1`}>{line.style}</span>}
            <span className="bg-emerald-100 text-emerald-900">
              <Runs runs={line.runs} />
            </span>
          </>
        )}
      </div>
    </div>
  );
}

export default function SideDocx({ lines }: { lines: DocLine[] }) {
  return (
    <div className="overflow-hidden rounded-lg border border-zinc-200 bg-white">
      {lines.map((l, i) => (
        <Line key={i} line={l} />
      ))}
    </div>
  );
}
