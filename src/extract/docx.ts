/**
 * DOCX extractor — port of extractors/docx.py: unzip, parse
 * word/document.xml, walk the body into ordered blocks. Tracked-changes
 * files are REFUSED (day-2 rule): a revision-marked document has no single
 * "after" state, so diffing would silently pick one interpretation.
 */

import JSZip from "jszip";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCUMENT_PART = "word/document.xml";
const REVISION_MARKERS = ["ins", "del", "moveFrom", "moveTo"];

export class DocxTrackedChangesError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "DocxTrackedChangesError";
  }
}

export interface Run {
  text: string;
  bold: boolean;
  italic: boolean;
}

export interface ParagraphBlock {
  blockType: "paragraph";
  style: string;
  runs: Run[];
}

export interface TableBlock {
  blockType: "table";
  rows: string[][];
}

export type Block = ParagraphBlock | TableBlock;

export interface Docx {
  blocks: Block[];
}

export const paragraphText = (p: ParagraphBlock): string => p.runs.map((r) => r.text).join("");

const direct = (el: Element, name: string) =>
  [...el.children].filter((c) => c.namespaceURI === W_NS && c.localName === name);

// <w:b/> means on; w:val="0"/"false" turns back off
function on(el: Element | undefined): boolean {
  if (!el) return false;
  const val = el.getAttributeNS(W_NS, "val") ?? "1";
  return val !== "0" && val !== "false";
}

// w:t / w:tab / w:br are real characters; w:delText never matches w:t, so
// deleted runs contribute nothing even if the refusal were skipped
function runText(r: Element): string {
  let out = "";
  for (const el of r.getElementsByTagName("*")) {
    if (el.namespaceURI !== W_NS) continue;
    if (el.localName === "t") out += el.textContent ?? "";
    else if (el.localName === "tab") out += "\t";
    else if (el.localName === "br" || el.localName === "cr") out += "\n";
  }
  return out;
}

function paragraph(p: Element): ParagraphBlock {
  let style = "Normal";
  const ppr = direct(p, "pPr")[0];
  const val = ppr ? direct(ppr, "pStyle")[0]?.getAttributeNS(W_NS, "val") : undefined;
  if (val) style = val;
  const runs: Run[] = [];
  for (const r of p.getElementsByTagNameNS(W_NS, "r")) {
    const text = runText(r);
    if (!text) continue;
    const rpr = direct(r, "rPr")[0];
    runs.push({
      text,
      bold: !!rpr && on(direct(rpr, "b")[0]),
      italic: !!rpr && on(direct(rpr, "i")[0]),
    });
  }
  return { blockType: "paragraph", style, runs };
}

// a cell is itself a block container: walk like the body so nested tables
// and sdt wrappers contribute (engine known-issue-2 lesson)
function cellText(tc: Element): string {
  const blocks: Block[] = [];
  walkInto(tc, blocks);
  const parts: string[] = [];
  for (const b of blocks) {
    if (b.blockType === "paragraph") {
      const t = paragraphText(b);
      if (t) parts.push(t);
    } else {
      for (const row of b.rows) parts.push(...row.filter((c) => c));
    }
  }
  return parts.join(" ");
}

function table(tbl: Element): TableBlock {
  return { blockType: "table", rows: direct(tbl, "tr").map((tr) => direct(tr, "tc").map(cellText)) };
}

function walkInto(parent: Element, blocks: Block[]): void {
  for (const child of parent.children) {
    if (child.namespaceURI !== W_NS) continue;
    if (child.localName === "p") blocks.push(paragraph(child));
    else if (child.localName === "tbl") blocks.push(table(child));
    else if (child.localName === "sdt" || child.localName === "sdtContent") walkInto(child, blocks);
    // w:sectPr & friends: layout metadata -> noise, skipped
  }
}

export async function extractDocx(
  bytes: Buffer | ArrayBuffer | Uint8Array,
  display = "document",
): Promise<Docx> {
  const zip = await JSZip.loadAsync(bytes);
  const part = zip.file(DOCUMENT_PART);
  if (!part) throw new Error(`${display}: not a .docx (no ${DOCUMENT_PART} inside)`);
  const doc = new DOMParser().parseFromString(await part.async("string"), "text/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
    throw new Error(`${display}: could not parse ${DOCUMENT_PART}`);
  }
  const root = doc.documentElement;
  for (const marker of REVISION_MARKERS) {
    if (root.getElementsByTagNameNS(W_NS, marker).length > 0) {
      throw new DocxTrackedChangesError(
        `${display}: tracked changes present; accept or reject them in Word first`,
      );
    }
  }
  const blocks: Block[] = [];
  const body = direct(root, "body")[0];
  if (body) walkInto(body, blocks);
  return { blocks };
}
