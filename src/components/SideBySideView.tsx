"use client";

/** Git-style side-by-side: both documents rendered in place with changes
 *  highlighted in context. Model logic lives in core/sideView.ts; this just
 *  picks the format and lays out the before/after header. */

import SideDocx from "./SideDocx";
import SideXlsx from "./SideXlsx";
import type { Change } from "../core/model";
import { sideDocx, sideXlsx, type DocPair } from "../core/sideView";

const MAX_DOCX_LINES = 400;

export default function SideBySideView({
  pair,
  changes,
  oldName,
  newName,
}: {
  pair: DocPair;
  changes: Change[];
  oldName?: string;
  newName?: string;
}) {
  const lines = pair.kind === "docx" ? sideDocx(pair.old, pair.new) : null;
  const shown = lines?.slice(0, MAX_DOCX_LINES) ?? null;
  return (
    <div className="mt-3">
      <div className="grid grid-cols-2 gap-4 text-[11px] font-medium text-zinc-500">
        <span className="truncate">before · {oldName}</span>
        <span className="truncate">after · {newName}</span>
      </div>

      {pair.kind === "xlsx" ? (
        <SideXlsx panes={sideXlsx(changes, pair.old, pair.new)} />
      ) : (
        <div className="mt-3">
          {shown && <SideDocx lines={shown} />}
          {lines && shown && lines.length > shown.length && (
            <div className="mt-1 rounded-md bg-zinc-50 px-3 py-2 text-xs text-zinc-500">
              + {lines.length - shown.length} more lines — see the change list
            </div>
          )}
        </div>
      )}

      <div className="mt-2 flex flex-wrap gap-3 text-[11px] text-zinc-500">
        <span className="rounded bg-red-100 px-2 py-0.5 text-red-900">before</span>
        <span className="rounded bg-emerald-100 px-2 py-0.5 text-emerald-900">after</span>
        <span>context is dimmed · an empty side means the block or cell does not exist there</span>
      </div>
    </div>
  );
}
