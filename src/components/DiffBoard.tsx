"use client";

import { useCallback, useRef, useState, useSyncExternalStore } from "react";

import ResultsView from "./ResultsView";
import { diffDocx } from "../core/docxDiff";
import { hashBytes } from "../core/hash";
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

const DEMO_BTN = "rounded-md border border-zinc-300 bg-white font-medium text-zinc-700 disabled:opacity-40";

// localStorage as a real store: SSR-safe, no setState-in-effect, tabs stay in sync
const recalcPrefSubscribe = (cb: () => void) => {
  window.addEventListener("storage", cb);
  window.addEventListener("anydiff-recalc", cb);
  return () => {
    window.removeEventListener("storage", cb);
    window.removeEventListener("anydiff-recalc", cb);
  };
};
const readRecalcPref = () => {
  try {
    return localStorage.getItem("anydiff.recalc") === "1";
  } catch {
    return false;
  }
};
const writeRecalcPref = (on: boolean) => {
  try {
    localStorage.setItem("anydiff.recalc", on ? "1" : "0");
  } catch {
    /* private mode: session-only */
  }
  window.dispatchEvent(new Event("anydiff-recalc"));
};

export default function DiffBoard() {
  const [oldFile, setOld] = useState<Slot>(null);
  const [newFile, setNew] = useState<Slot>(null);
  const [changes, setChanges] = useState<Change[] | null>(null);
  const [pair, setPair] = useState<DocPair | null>(null);
  const [mode, setMode] = useState<"side" | "list">("side");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const recalc = useSyncExternalStore(recalcPrefSubscribe, readRecalcPref, () => false);
  const [recalcNote, setRecalcNote] = useState<string | null>(null);
  const [computedMode, setComputedMode] = useState(false);

  const run = useCallback(async (old: NonNullable<Slot>, fresh: NonNullable<Slot>, recalcOn: boolean) => {
    setBusy(true);
    setError(null);
    setRecalcNote(null);
    setComputedMode(false);
    try {
      if (old.ext !== fresh.ext) {
        throw new Error("Both files must be the same type — .xlsx with .xlsx, .docx with .docx.");
      }
      const [a, b] = await Promise.all([old.bytes, fresh.bytes]);
      if (a.byteLength === b.byteLength && hashBytes(new Uint8Array(a)) === hashBytes(new Uint8Array(b))) {
        // same bytes: nothing a differ could report — extract once, zero changes
        if (old.ext === "docx") {
          const d = await extractDocx(a, old.name);
          setPair({ kind: "docx", old: d, new: d });
        } else {
          const s = await extractXlsx(a);
          setPair({ kind: "xlsx", old: s, new: s });
        }
        setChanges([]);
        return;
      }
      if (old.ext === "docx") {
        const d1 = await extractDocx(a, old.name);
        const d2 = await extractDocx(b, fresh.name);
        setPair({ kind: "docx", old: d1, new: d2 });
        setChanges(diffDocx(d1, d2, fresh.name));
      } else {
        let s1 = await extractXlsx(a);
        let s2 = await extractXlsx(b);
        if (recalcOn) {
          const { recalcSheets } = await import("../core/recalc"); // engine chunk loads only on demand
          const r1 = recalcSheets(s1);
          const r2 = recalcSheets(s2);
          const skip = r1.skipped ?? r2.skipped;
          if (skip) {
            setRecalcNote(`recalculation skipped: ${skip}`);
          } else {
            s1 = r1.sheets;
            s2 = r2.sheets;
            setComputedMode(true);
            const n = r1.computed + r2.computed;
            const f = r1.failed + r2.failed;
            setRecalcNote(
              `${n} formula cells recomputed in the browser${f ? `, ${f} kept as saved (errors, dates, exotic results)` : ""} — values below are computed, not what the files stored`,
            );
          }
        }
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
    await run(old, fresh, recalc);
  }

  const bothReady = oldFile && newFile;
  const fullPage = mode === "side" && !!pair && !!changes && changes.length > 0;

  return (
    <main className="w-full px-6 py-12">
      {/* the input bar never moves — it keeps the change-list width in both modes */}
      <div className="mx-auto w-full max-w-5xl">
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
            onClick={() => oldFile && newFile && run(oldFile, newFile, recalc)}
            className="rounded-md bg-zinc-900 px-4 py-2 text-sm font-medium text-white disabled:opacity-40"
          >
            {busy ? "Diffing…" : "Compare"}
          </button>
          <button
            type="button"
            disabled={busy}
            onClick={() => loadDemo("xlsx")}
            className={`${DEMO_BTN} px-4 py-2 text-sm`}
          >
            Q3 spreadsheet demo
          </button>
          <button
            type="button"
            disabled={busy}
            onClick={() => loadDemo("docx")}
            className={`${DEMO_BTN} px-4 py-2 text-sm`}
          >
            Meeting notes demo
          </button>
          <label
            className={`ml-auto flex select-none items-center gap-2 text-xs ${oldFile?.ext === "docx" || newFile?.ext === "docx" ? "text-zinc-300" : "text-zinc-600"}`}
            title="xlsx only: evaluate every formula in the browser and diff the computed results instead of the values the files stored"
          >
            <input
              type="checkbox"
              className="size-3.5 accent-sky-600"
              disabled={oldFile?.ext === "docx" || newFile?.ext === "docx"}
              checked={recalc}
              onChange={(e) => writeRecalcPref(e.target.checked)}
            />
            recalculate formulas in browser
          </label>
        </div>

        {error && <p className="mt-4 rounded-md bg-red-50 p-3 text-sm text-red-800">{error}</p>}
      </div>

      <div className={fullPage ? "" : "mx-auto w-full max-w-5xl"}>
        {changes && recalcNote && (
          <div className="mt-3 rounded-md border border-sky-200 bg-sky-50 px-3 py-1.5 text-xs text-sky-900">
            computed mode · {recalcNote}
          </div>
        )}
        {changes && (
          <ResultsView
            key={`${oldFile?.name}->${newFile?.name}`}
            changes={changes}
            pair={pair}
            mode={mode}
            onMode={setMode}
            oldName={oldFile?.name}
            newName={newFile?.name}
            computed={computedMode}
          />
        )}
      </div>
    </main>
  );
}
