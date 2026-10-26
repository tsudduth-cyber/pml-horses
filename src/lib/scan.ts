import { graph, graphRaw } from "./graph";
import { CARE_TASKS } from "./types";
import type { CareKey, Doc, Horse } from "./types";

const DRIVE = process.env.SP_DRIVE_ID!;
const ROOT = process.env.SP_HORSES_PATH!;

/** Where the scan result lives. One small file, read on every page load. */
const SCAN_PATH = () => `${ROOT}/_app/scan.json`;

export interface ScanHit {
  date: string;
  /** The file it was read out of. */
  from: string;
  /** The line of the document it came from, so a wrong date is traceable. */
  line: string;
}

export interface Scan {
  at: string;
  /** horse id → care key → what the paperwork says. */
  items: Record<string, Partial<Record<CareKey, ScanHit>>>;
}

/* ---------- reading the saved scan ---------- */

export async function readScan(): Promise<Scan | null> {
  try {
    const path = encodeURIComponent(SCAN_PATH()).replace(/%2F/g, "/");
    const res = await graphRaw(`/drives/${DRIVE}/root:/${path}:/content`);
    const data = JSON.parse(await res.text());
    if (data && typeof data.at === "string" && data.items) return data as Scan;
  } catch {
    /* no scan yet, or it is unreadable — treat as none */
  }
  return null;
}

async function writeScan(scan: Scan) {
  const path = encodeURIComponent(SCAN_PATH()).replace(/%2F/g, "/");
  await graph(`/drives/${DRIVE}/root:/${path}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(scan, null, 1),
  });
}

/** How old a scan may get before the next person to open the app redoes it. */
export const STALE_MS = 4 * 60 * 60 * 1000;

export function scanIsStale(scan: Scan | null) {
  if (!scan) return true;
  const at = Date.parse(scan.at);
  return !Number.isFinite(at) || Date.now() - at > STALE_MS;
}

/* ---------- reading what a document says ---------- */

const MONTHS = "jan feb mar apr may jun jul aug sep oct nov dec".split(" ");

/**
 * Every date in a block of text, with where it sits.
 *
 * Invoices write dates out long ("Oct 13, 2026"); filenames use digits. Both
 * are here because the same parser reads both.
 */
function datesIn(text: string): { at: number; date: string }[] {
  const out: { at: number; date: string }[] = [];

  const long = /\b([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})\b/g;
  for (const m of text.matchAll(long)) {
    const mi = MONTHS.indexOf(m[1].slice(0, 3).toLowerCase());
    if (mi < 0) continue;
    const d = iso(+m[3], mi + 1, +m[2]);
    if (d) out.push({ at: m.index!, date: d });
  }

  const num = /\b(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})\b/g;
  for (const m of text.matchAll(num)) {
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    const d = iso(y, +m[1], +m[2]);
    if (d) out.push({ at: m.index!, date: d });
  }

  return out.sort((a, b) => a.at - b.at);
}

function iso(y: number, m: number, d: number): string | null {
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 2000) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  if (dt.getTime() > Date.now() + 864e5) return null; // nothing in the future
  return dt.toISOString().slice(0, 10);
}

const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Every mention of a horse by barn or registered name, with where it sits. */
function horsesIn(text: string, horses: Horse[]): { at: number; id: string }[] {
  const out: { at: number; id: string }[] = [];
  for (const h of horses) {
    for (const name of [h.barn, h.reg]) {
      if (!name || name.length < 2) continue;
      const re = new RegExp(`(^|[^a-z0-9])${esc(name)}([^a-z0-9]|$)`, "gi");
      for (const m of text.matchAll(re)) out.push({ at: m.index!, id: h.id });
    }
  }
  return out.sort((a, b) => a.at - b.at);
}

/** The nearest entry at or before `at`, within `window` characters. */
function nearest<T extends { at: number }>(
  list: T[],
  at: number,
  window: number
): T | null {
  let best: T | null = null;
  for (const e of list) {
    if (e.at > at) break;
    if (at - e.at <= window) best = e;
  }
  return best;
}

/**
 * Pull care events out of a document's text.
 *
 * A vet invoice is a table — patient, date, item — flattened into one run of
 * text. So for each treatment word, take the horse and the date that most
 * recently preceded it. A line that does not have both is skipped: the point
 * of reading the file is a better answer than the filename, not more guesses.
 */
export function eventsIn(
  text: string,
  horses: Horse[]
): { id: string; key: CareKey; date: string; line: string }[] {
  const dates = datesIn(text);
  const named = horsesIn(text, horses);
  const out: { id: string; key: CareKey; date: string; line: string }[] = [];

  for (const t of CARE_TASKS) {
    const re = new RegExp(t.match.source, "gi");
    for (const m of text.matchAll(re)) {
      const d = nearest(dates, m.index!, 160);
      const h = nearest(named, m.index!, 160);
      if (!d || !h) continue;
      out.push({
        id: h.id,
        key: t.key,
        date: d.date,
        line: text.slice(Math.max(0, m.index! - 70), m.index! + 40).trim(),
      });
    }
  }
  return out;
}

/* ---------- the scan itself ---------- */

async function pdfText(itemId: string): Promise<string> {
  const res = await graphRaw(`/drives/${DRIVE}/items/${itemId}/content`);
  const buf = new Uint8Array(await res.arrayBuffer());
  const { extractText, getDocumentProxy } = await import("unpdf");
  const doc = await getDocumentProxy(buf);
  const { text } = await extractText(doc, { mergePages: true });
  return typeof text === "string" ? text : (text as string[]).join("\n");
}

/**
 * Read every PDF once and record what it says.
 *
 * Run from the app by whoever opens it when the last scan has gone stale —
 * not on a schedule. Graph access here is delegated, so a cron job would
 * have nobody to act as; the person opening the app does.
 */
export async function runScan(
  horses: Horse[],
  docs: { doc: Doc; ids: string[] }[]
): Promise<Scan> {
  const items: Scan["items"] = {};

  const pdfs = docs
    .filter((d) => !d.doc.folder && /\.pdf$/i.test(d.doc.name) && d.doc.id)
    .sort((a, b) => (b.doc.modified ?? "").localeCompare(a.doc.modified ?? ""))
    .slice(0, 120); // newest first, bounded so one bad folder cannot hang it

  await Promise.all(
    pdfs.map(async (entry) => {
      let text = "";
      try {
        text = await pdfText(entry.doc.id);
      } catch {
        return; // a scan, a locked file, an odd PDF — skip it quietly
      }
      if (text.trim().length < 40) return; // an image-only scan says nothing

      // A document found through a horse's own folder is about that horse
      // even when the text never names her.
      const scoped = horses.filter((h) => entry.ids.includes(h.id));
      const found = eventsIn(text, horses.length ? horses : []);

      for (const ev of found) {
        // Only credit horses this document is already associated with, when
        // it came from a specific horse's folder.
        if (entry.ids.length && !entry.ids.includes(ev.id)) continue;
        const bucket = (items[ev.id] ??= {});
        const held = bucket[ev.key];
        if (!held || ev.date > held.date) {
          bucket[ev.key] = { date: ev.date, from: entry.doc.name, line: ev.line };
        }
      }

      // Nothing named a horse, but the folder did — still useful.
      if (!found.length && scoped.length === 1) return;
    })
  );

  const scan: Scan = { at: new Date().toISOString(), items };
  try {
    await writeScan(scan);
  } catch {
    /* the scan still returns; it just will not be reused */
  }
  return scan;
}
