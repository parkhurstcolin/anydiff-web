"use client";

import { useCallback, useRef, useState } from "react";

import { diffXlsx } from "../core/xlsxDiff";
import { fmtValue, visible } from "../core/format";
import type { Change } from "../core/model";
import { extractXlsx } from "../extract/xlsx";

const CAP = 40;

type Slot = { name: string; bytes: Promise<ArrayBuffer> } | null;

const KIND_STYLE: Record<string, string> = {
  cell: "bg-zinc-100 text-zinc-600",
  sheet_added: "bg-emerald-100 text-emerald-800",
  sheet_removed: "bg-red-100 text-red-800",
  sheet_renamed: "bg-amber-100 text-amber-800",
};

function disp(v: string | number | boolean | null): string {
  if (v === null) return "—";
  const s = fmtValue(v);
  return typeof v === "string" ? `"${visible(s)}"` : s;
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
      className={`flex-1 rounded-lg border-2 border-dashed p-6 text-left transition ${
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

  const run = useCallback(async (old: NonNullable<Slot>, fresh: NonNullable<Slot>) => {
    setBusy(true);
    setError(null);
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

  const pick = (setter: (s: Slot) => void) => (name: string, bytes: Promise<ArrayBuffer>) => {
    setter({ name, bytes });
    return true;
  };

  async function loadDemo() {
    const fetch2 = async (p: string) => {
      const r = await fetch(p);
      if (!r.ok) throw new Error(`demo file ${p} failed: ${r.status}`);
      return r.arrayBuffer();
    };
    const old: NonNullable<Slot> = { name: "Q3-forecast.xlsx", bytes: fetch2("/demo/Q3-forecast.xlsx") };
    const fresh: NonNullable<Slot> = { name: "Q3-forecast-v2.xlsx", bytes: fetch2("/demo/Q3-forecast-v2.xlsx") };
    setOld(old);
    setNew(fresh);
    await run(old, fresh);
  }

  const bothReady = oldFile && newFile;

  return (
    <main className="mx-auto w-full max-w-4xl px-6 py-12">
      <h1 className="text-2xl font-semibold tracking-tight text-zinc-900">AnyDiff</h1>
      <p className="mt-1 text-sm text-zinc-600">
        See what actually changed between two Excel files — values, formulas, renamed sheets.
        Everything runs in your browser; your files never leave this tab.
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

      {changes && (
        <section className="mt-6">
          <h2 className="text-sm font-medium text-zinc-900">
            {changes.length} change{changes.length === 1 ? "" : "s"}
            {changes.length === 0 && " — the files are semantically identical"}
          </h2>
          {changes.length > 0 && (
            <table className="mt-3 w-full border-collapse text-sm">
              <thead>
                <tr className="border-b border-zinc-200 text-left text-xs uppercase tracking-wide text-zinc-500">
                  <th className="py-2 pr-3 font-medium">where</th>
                  <th className="py-2 pr-3 font-medium">before → after</th>
                  <th className="py-2 font-medium">note</th>
                </tr>
              </thead>
              <tbody>
                {changes.slice(0, CAP).map((c, i) => (
                  <tr key={i} className="border-b border-zinc-100 align-top">
                    <td className="py-2 pr-3">
                      <span className={`rounded px-1.5 py-0.5 text-xs font-medium ${KIND_STYLE[c.kind] ?? ""}`}>
                        {c.kind === "cell" ? "cell" : c.kind.replace("sheet_", "")}
                      </span>{" "}
                      <code className="text-xs">{c.where}</code>
                      {c.confidence !== "exact" && (
                        <span className="ml-1 rounded bg-amber-100 px-1.5 py-0.5 text-xs text-amber-800">
                          {c.confidence === "moved" ? "moved" : "moved, inspect manually"}
                        </span>
                      )}
                    </td>
                    <td className="py-2 pr-3 font-mono text-xs">
                      {disp(c.before)} <span className="text-zinc-400">→</span> {disp(c.after)}
                    </td>
                    <td className="py-2 text-xs text-zinc-600">{c.note ?? ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
          {changes.length > CAP && (
            <p className="mt-2 text-xs text-zinc-500">
              + {changes.length - CAP} more — full report coming with HTML export
            </p>
          )}
        </section>
      )}
    </main>
  );
}
