"use client";

import type { GridRow, SheetPane } from "../core/sideView";

const SIDE = "bg-emerald-100 text-emerald-900";
const LEFT = "bg-red-100 text-red-900";
const CHIP = "rounded bg-amber-100 px-1 py-0.5 text-[10px] font-medium text-amber-900";
const GUTTER = "text-[10px] text-zinc-400";

const colLetter = (n: number): string => {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

function Grid({ rows, col0, side }: { rows: GridRow[]; col0: number; side: "old" | "new" }) {
  const changed = side === "old" ? LEFT : SIDE;
  const letters: string[] = [];
  for (let c = col0; c < col0 + (rows[0]?.cells.length ?? 0); c++) letters.push(colLetter(c));
  return (
    <div className="max-h-[75vh] overflow-auto rounded border border-zinc-200">
      <table className="border-collapse">
        <thead>
          <tr className="text-[10px] text-zinc-400">
            <th className="sticky left-0 top-0 z-20 border-b border-r border-zinc-200 bg-zinc-50 px-1.5 py-1" />
            {letters.map((l) => (
              <th key={l} className="sticky top-0 z-10 min-w-16 border-b border-r border-zinc-200 bg-zinc-50 px-2 py-1 font-medium">
                {l}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.n}>
              <th className="sticky left-0 z-10 border-b border-r border-zinc-200 bg-zinc-50 px-1.5 py-0.5 text-right align-top text-[10px] font-medium text-zinc-400">
                {r.n}
              </th>
              {r.cells.map((cell, i) => (
                <td
                  key={i}
                  title={cell?.title ?? undefined}
                  className={`max-w-56 truncate border-b border-r border-zinc-100 px-2 py-0.5 font-mono text-xs ${
                    cell?.changed ? changed : "text-zinc-700"
                  }`}
                >
                  {cell?.text ?? ""}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Pane({ pane }: { pane: SheetPane }) {
  const label = (s: string) => <span className="truncate font-mono text-xs text-zinc-700">{s || "—"}</span>;
  return (
    <div className="mt-4 overflow-hidden rounded-lg border border-zinc-200 bg-white">
      <div className="flex flex-wrap items-center gap-2 border-b border-zinc-200 bg-zinc-50 px-3 py-1.5">
        {pane.rename ? (
          <span className="text-xs text-zinc-700">
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
        {pane.note && <span className={CHIP}>{pane.note}</span>}
        {pane.truncated && <span className={GUTTER}>window around changes</span>}
      </div>
      <div className="grid grid-cols-2 gap-4 p-3">
        {pane.kind === "added" ? <div className={GUTTER}>not in before</div> : <Grid rows={pane.oldRows} col0={pane.col0} side="old" />}
        {pane.kind === "removed" ? <div className={GUTTER}>not in after</div> : <Grid rows={pane.newRows} col0={pane.col0} side="new" />}
      </div>
    </div>
  );
}

export default function SideXlsx({ panes }: { panes: SheetPane[] }) {
  return (
    <>
      {panes.map((p, i) => (
        <Pane key={i} pane={p} />
      ))}
    </>
  );
}
