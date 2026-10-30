"use client";

import { useCallback, useRef, useState } from "react";

import { diffXlsx } from "../core/xlsxDiff";
import { fmtValue, visible } from "../core/format";
import type { Change } from "../core/model";
import { diffWords } from "../core/wordDiff";
import { extractXlsx } from "../extract/xlsx";

const CAP = 40;

type Slot = { name: string; bytes: Promise<ArrayBuffer> } | null;

const ADDED = "rounded bg-emerald-100 px-0.5 text-emerald-900";
const REMOVED = "rounded bg-red-100 px-0.5 text-red-900 line-through decoration-red-400";

function CellValueDiff({ c }: { c: Change }) {
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

function rowLabel(c: Change): { badge: string; badgeStyle: string } {
  if (c.kind !== "cell") {
    return {
      badge: c.kind.replace("sheet_", ""),
      badgeStyle: c.kind === "sheet_added" ? ADDED : c.kind === "sheet_removed" ? REMOVED : "rounded bg-amber-100 px-1 text-amber-900",
    };
  }
  if (c.before === null) return { badge: "added", badgeStyle: ADDED };
  if (c.after === null) return { badge: "removed", badgeStyle: REMOVED };
  return { badge: "edit", badgeStyle: "rounded bg-zinc-100 px-1 text-zinc-600" };
}

function DropZone({
  label,
  slot,
  onPick,
}: {
  label: string;
  slot: Slot;
  onPick: (name: string, bytes: Promise<ArrayBuffer>) => void;
}) {
  const input = useRef<HTMLInputElement>(null);
  return (
    <button
      type="button"
      onClick={() => input.current?.click()}
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        e.preventDefault();
        const f = e.dataTransfer.files?.[0];
        if (f) onPick(f.name, f.arrayBuffer());
      }}
      className={`flex-1 rounded-lg border-2 border-dashed p-5 text-left transition ${
        slot ? "border-sky-400 bg-sky-50" : "border-zinc-300 bg-white hover:border-zinc-400"
      }`}
    >
      <span className="text-xs font-medium uppercase tracking-wide text-zinc-500">{label}</span>
      <p className="mt-1 truncate text-sm font-medium text-zinc-900">
        {slot ? slot.name : "drop .xlsx or click to choose"}
      </p>
      <input
        ref={input}
        type="file"
        accept=".xlsx"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onPick(f.name, f.arrayBuffer());
        }}
      />
    </button>
  );
}

export default function DiffBoard() {
  const [oldFile, setOld] = useState<Slot>(null);
  const [newFile, setNew] = useState<Slot>(null);
  const [changes, setChanges] = useState<Change[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [showAll, setShowAll] = useState(false);

  const run = useCallback(async (old: NonNullable<Slot>, fresh: NonNullable<Slot>) => {
    setBusy(true);
    setError(null);
    setShowAll(false);
    try {
      const [a, b] = await Promise.all([old.bytes, fresh.bytes]);
      setChanges(diffXlsx(await extractXlsx(a), await extractXlsx(b), fresh.name));
    } catch (e) {
      setChanges(null);
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, []);

  const pick = (setter: (s: Slot) => void) => (name: string, bytes: Promise<ArrayBuffer>) =>
    setter({ name, bytes });

  async function loadDemo() {
    const f = async (p: string) => {
      const r = await fetch(p);
      if (!r.ok) throw new Error(`demo file ${p} failed: ${r.status}`);
      return r.arrayBuffer();
    };
    const old: NonNullable<Slot> = { name: "Q3-forecast.xlsx", bytes: f("/demo/Q3-forecast.xlsx") };
    const fresh: NonNullable<Slot> = { name: "Q3-forecast-v2.xlsx", bytes: f("/demo/Q3-forecast-v2.xlsx") };
    setOld(old);
    setNew(fresh);
    await run(old, fresh);
  }

  const bothReady = oldFile && newFile;
  const stats = changes
    ? {
        edits: changes.filter((c) => c.kind === "cell" && c.before !== null && c.after !== null).length,
        adds: changes.filter((c) => c.kind === "cell" && c.before === null).length,
        removes: changes.filter((c) => c.kind === "cell" && c.after === null).length,
        sheets: changes.filter((c) => c.kind !== "cell").length,
      }
    : null;
  const structural = changes?.filter((c) => c.kind !== "cell") ?? [];
  const cells = changes?.filter((c) => c.kind === "cell") ?? [];
  const bySheet = new Map<string, Change[]>();
  for (const c of cells) {
    const sheet = c.where.includes("!") ? c.where.slice(0, c.where.indexOf("!")) : "—";
    (bySheet.get(sheet) ?? bySheet.set(sheet, []).get(sheet)!).push(c);
  }

  return (
    <main className="mx-auto w-full max-w-5xl px-6 py-12">
      <h1 className="text-2xl font-semibold tracking-tight text-zinc-900">AnyDiff</h1>
      <p className="mt-1 max-w-xl text-sm text-zinc-600">
        What actually changed between two Excel files — values, formulas, renamed sheets.
        Everything runs in your browser; files never leave this tab.
      </p>

      <div className="mt-6 flex gap-4">
        <DropZone label="before" slot={oldFile} onPick={pick(setOld)} />
        <DropZone label="after" slot={newFile} onPick={pick(setNew)} />
      </div>

      <div className="mt-4 flex items-center gap-3">
        <button
          type="button"
          disabled={!bothReady || busy}
          onClick={() => bothReady && run(oldFile, newFile)}
          className="rounded-md bg-zinc-900 px-4 py-2 text-sm font-medium text-white disabled:opacity-40"
        >
          {busy ? "Diffing…" : "Compare"}
        </button>
        <button
          type="button"
          disabled={busy}
          onClick={loadDemo}
          className="rounded-md border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-700 disabled:opacity-40"
        >
          Try the Q3 demo
        </button>
        <span className="text-xs text-zinc-500">.docx support: next milestone</span>
      </div>

      {error && <p className="mt-4 rounded-md bg-red-50 p-3 text-sm text-red-800">{error}</p>}

      {changes && stats && (
        <section className="mt-8">
          <div className="flex flex-wrap items-baseline gap-x-4 gap-y-1">
            <h2 className="text-sm font-semibold text-zinc-900">
              {changes.length === 0
                ? "Semantically identical"
                : `${stats.edits} edit${stats.edits === 1 ? "" : "s"} · ${stats.adds} added · ${stats.removes} removed${
                    stats.sheets ? ` · ${stats.sheets} sheet-level` : ""
                  }`}
            </h2>
            <span className="text-xs text-zinc-500">
              {oldFile?.name} → {newFile?.name}
            </span>
          </div>

          {changes.length === 0 && (
            <p className="mt-3 rounded-md bg-emerald-50 p-3 text-sm text-emerald-800">
              Every sheet, value and formula matches. Formatting-only edits are deliberately invisible.
            </p>
          )}

          {structural.length > 0 && (
            <div className="mt-4 rounded-lg border border-zinc-200 bg-white">
              {structural.map((c, i) => {
                const l = rowLabel(c);
                return (
                  <div key={i} className="flex items-center gap-3 border-b border-zinc-100 px-4 py-2 text-sm last:border-b-0">
                    <span className={`text-xs font-medium ${l.badgeStyle}`}>{l.badge}</span>
                    <span className="font-mono text-xs">{c.where}</span>
                    {c.confidence !== "exact" && (
                      <span className="rounded bg-amber-100 px-1.5 py-0.5 text-xs text-amber-900">
                        {c.confidence === "moved" ? "moved" : "moved, inspect manually"}
                      </span>
                    )}
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

          {[...bySheet.entries()].map(([sheet, rows]) => {
            const shown = showAll ? rows : rows.slice(0, CAP);
            return (
              <div key={sheet} className="mt-4 overflow-hidden rounded-lg border border-zinc-200 bg-white">
                <div className="border-b border-zinc-200 bg-zinc-50 px-4 py-2 text-xs font-semibold text-zinc-700">
                  {sheet} <span className="font-normal text-zinc-500">· {rows.length} change{rows.length === 1 ? "" : "s"}</span>
                </div>
                <table className="w-full border-collapse text-sm">
                  <tbody>
                    {shown.map((c, i) => {
                      const l = rowLabel(c);
                      return (
                        <tr key={i} className="border-b border-zinc-100 align-top last:border-b-0">
                          <td className="w-24 px-4 py-2">
                            <span className={`text-xs font-medium ${l.badgeStyle}`}>{l.badge}</span>
                          </td>
                          <td className="w-32 py-2 pr-3 font-mono text-xs text-zinc-700">
                            {c.where.includes("!") ? c.where.slice(c.where.indexOf("!") + 1) : c.where}
                            {c.confidence !== "exact" && (
                              <div className="mt-1 rounded bg-amber-100 px-1 py-0.5 text-[10px] text-amber-900">
                                {c.confidence === "moved" ? "moved" : "inspect manually"}
                              </div>
                            )}
                          </td>
                          <td className="py-2 pr-4 font-mono text-xs">
                            <CellValueDiff c={c} />
                            {c.note && <div className="mt-1 text-[11px] not-italic text-zinc-500">{c.note}</div>}
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
        </section>
      )}
    </main>
  );
}
