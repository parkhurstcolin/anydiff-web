"use client";

import { useCallback, useRef, useState } from "react";

import ResultsView from "./ResultsView";
import { diffDocx } from "../core/docxDiff";
import { diffXlsx } from "../core/xlsxDiff";
import type { Change } from "../core/model";
import type { DocPair } from "../core/sideView";
import { extractDocx } from "../extract/docx";
import { extractXlsx } from "../extract/xlsx";

type Ext = "xlsx" | "docx";
type Slot = { name: string; ext: Ext; bytes: Promise<ArrayBuffer> } | null;

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
        {slot ? slot.name : "drop .xlsx / .docx or click to choose"}
      </p>
      <input
        ref={input}
        type="file"
        accept=".xlsx,.docx"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onPick(f.name, f.arrayBuffer());
        }}
      />
    </button>
  );
}

const DEMO_BTN =
  "rounded-md border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-700 disabled:opacity-40";

export default function DiffBoard() {
  const [oldFile, setOld] = useState<Slot>(null);
  const [newFile, setNew] = useState<Slot>(null);
  const [changes, setChanges] = useState<Change[] | null>(null);
  const [pair, setPair] = useState<DocPair | null>(null);
  const [mode, setMode] = useState<"side" | "list">("side");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const run = useCallback(async (old: NonNullable<Slot>, fresh: NonNullable<Slot>) => {
    setBusy(true);
    setError(null);
    try {
      if (old.ext !== fresh.ext) {
        throw new Error("Both files must be the same type — .xlsx with .xlsx, .docx with .docx.");
      }
      const [a, b] = await Promise.all([old.bytes, fresh.bytes]);
      if (old.ext === "docx") {
        const d1 = await extractDocx(a, old.name);
        const d2 = await extractDocx(b, fresh.name);
        setPair({ kind: "docx", old: d1, new: d2 });
        setChanges(diffDocx(d1, d2, fresh.name));
      } else {
        const s1 = await extractXlsx(a);
        const s2 = await extractXlsx(b);
        setPair({ kind: "xlsx", old: s1, new: s2 });
        setChanges(diffXlsx(s1, s2, fresh.name));
      }
    } catch (e) {
      setChanges(null);
      setPair(null);
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, []);

  const pick = (setter: (s: Slot) => void) => (name: string, bytes: Promise<ArrayBuffer>) => {
    const lower = name.toLowerCase();
    const ext: Ext | null = lower.endsWith(".xlsx") ? "xlsx" : lower.endsWith(".docx") ? "docx" : null;
    if (!ext) {
      setError(`${name}: unsupported file — .xlsx and .docx only.`);
      return;
    }
    setError(null);
    setter({ name, ext, bytes });
  };

  async function loadDemo(kind: Ext) {
    const [oldName, newName] =
      kind === "xlsx"
        ? ["Q3-forecast.xlsx", "Q3-forecast-v2.xlsx"]
        : ["meeting-notes.docx", "meeting-notes-v2.docx"];
    const f = async (p: string) => {
      const r = await fetch(`/demo/${p}`);
      if (!r.ok) throw new Error(`demo file ${p} failed: ${r.status}`);
      return r.arrayBuffer();
    };
    const old: NonNullable<Slot> = { name: oldName, ext: kind, bytes: f(oldName) };
    const fresh: NonNullable<Slot> = { name: newName, ext: kind, bytes: f(newName) };
    setOld(old);
    setNew(fresh);
    await run(old, fresh);
  }

  const bothReady = oldFile && newFile;
  const fullPage = mode === "side" && !!pair && !!changes && changes.length > 0;

  return (
    <main className={`px-6 py-12 ${fullPage ? "w-full" : "mx-auto w-full max-w-5xl"}`}>
      <h1 className="text-2xl font-semibold tracking-tight text-zinc-900">AnyDiff</h1>
      <p className="mt-1 max-w-xl text-sm text-zinc-600">
        What actually changed between two Excel or Word files — values, formulas, paragraphs,
        tables, renamed sheets. Everything runs in your browser; files never leave this tab.
      </p>

      <div className="mt-6 flex gap-4">
        <DropZone label="before" slot={oldFile} onPick={pick(setOld)} />
        <DropZone label="after" slot={newFile} onPick={pick(setNew)} />
      </div>

      <div className="mt-4 flex flex-wrap items-center gap-3">
        <button
          type="button"
          disabled={!bothReady || busy}
          onClick={() => oldFile && newFile && run(oldFile, newFile)}
          className="rounded-md bg-zinc-900 px-4 py-2 text-sm font-medium text-white disabled:opacity-40"
        >
          {busy ? "Diffing…" : "Compare"}
        </button>
        <button type="button" disabled={busy} onClick={() => loadDemo("xlsx")} className={DEMO_BTN}>
          Q3 spreadsheet demo
        </button>
        <button type="button" disabled={busy} onClick={() => loadDemo("docx")} className={DEMO_BTN}>
          Meeting notes demo
        </button>
      </div>

      {error && <p className="mt-4 rounded-md bg-red-50 p-3 text-sm text-red-800">{error}</p>}

      {changes && (
        <ResultsView
          key={`${oldFile?.name}->${newFile?.name}`}
          changes={changes}
          pair={pair}
          mode={mode}
          onMode={setMode}
          oldName={oldFile?.name}
          newName={newFile?.name}
        />
      )}
    </main>
  );
}
