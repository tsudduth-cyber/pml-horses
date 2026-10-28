import { graph, graphRaw } from "./graph";
import { CARE_TASKS } from "./types";
import type { CareKey, Charge, Doc, Horse } from "./types";

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

/**
 * Bumped whenever the reading changes — a new folder comes into view, a
 * keyword is added, the matching is corrected. A saved scan from an older
 * version is wrong rather than merely old, so it is redone on the next
 * open instead of being trusted for another four hours.
 *
 * 2: read every folder beside /Horses/, credit a horse named on either
 *    side of the service word, count half sets and pads as farrier work.
 * 3: also read what each line cost, so a horse carries her own spend.
 * 4: read signed breeding contracts for money committed but not billed.
 * 5: only read charges from documents that say they are bills, so a
 *    balance sheet's carrying values stop reading as care costs.
 */
export const SCAN_VERSION = 5;

/** One charge read off an invoice, attributed to one horse. */
export type SpendHit = Charge;

export interface Scan {
  at: string;
  /** The reading that produced it. See SCAN_VERSION. */
  v?: number;
  /** horse id → care key → what the paperwork says. */
  items: Record<string, Partial<Record<CareKey, ScanHit>>>;
  /** horse id → every charge billed to her by name. */
  spend?: Record<string, SpendHit[]>;
  /** horse id → money committed in a signed contract but not yet billed. */
  committed?: Record<string, SpendHit[]>;
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
  if ((scan.v ?? 1) !== SCAN_VERSION) return true;
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
 * The nearest entry on either side of `at`, within `window` characters.
 *
 * Which side the horse's name falls on depends on how the invoice is laid
 * out. "Sully | Oct 13 | Coggins" puts it before the service; the farrier's
 * "10/26/2026 Trim lizard" puts it after. Looking only backwards credited
 * every line on that second kind to the horse on the line above, so the
 * whole barn came out shifted by one. Ties go to the earlier one, which is
 * the way a row usually reads.
 */
function nearestEither<T extends { at: number }>(
  list: T[],
  at: number,
  window: number
): T | null {
  let best: T | null = null;
  let gap = Infinity;
  for (const e of list) {
    const d = Math.abs(e.at - at);
    if (d > window) continue;
    if (d < gap) {
      gap = d;
      best = e;
    }
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
      const h = nearestEither(named, m.index!, 160);
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

/* ---------- reading what a document costs ---------- */

/**
 * A money amount: either it carries cents, or it carries a dollar sign.
 *
 * The lookarounds keep "10.26.2026" in a filename from reading as $10.26,
 * and a bare quantity column from reading as money.
 */
const MONEY = /(?<![\d.])\$\s?\d[\d,]*(?:\.\d{2})?(?![\d.])|(?<![\d.])\d[\d,]*\.\d{2}(?![\d.])/g;

/** A line that restates money already counted somewhere else. */
const SUMMARY_LINE =
  /\b(total|subtotal|balance|amount due|payment|paid|tax|deposit|credit|due|invoice amount)\b/i;

/**
 * A document has to say it is a bill before a number in it is treated as
 * a charge.
 *
 * This is an allow-list on purpose. A balance sheet lists every horse
 * beside her carrying value — "Vintage Enterprize (JP) 30,000.00" — which
 * reads exactly like an invoice row and is not a cost at all; three of
 * them on file would have put $90,000 of "care" on JP. Missing a real
 * invoice shows up as a horse with no charges, which is visible. Reading
 * a balance sheet as charges shows up as a plausible wrong number, which
 * is not.
 */
const LOOKS_LIKE_A_BILL =
  /\b(invoice|receipt|bill to|billed to|amount due|balance due|remit|statement of account)\b/i;

/** Documents that list money but bill nothing. */
const NOT_A_BILL =
  /\b(balance sheet|income statement|profit (and|&) loss|p&l|general ledger|trial balance|settlement sheet|agreement|contract|certificate|registration|transfer form)\b/i;

function toAmount(s: string): number {
  const n = Number(s.replace(/[$,\s]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

/** Every money amount in a string, in order. */
function moneyIn(s: string): number[] {
  return (s.match(MONEY) ?? []).map(toAmount).filter((n) => n > 0);
}

/**
 * What each horse was charged on one document.
 *
 * Worked line by line rather than by proximity, because an invoice is a
 * table and a row is the unit that means something — charging one horse
 * for the row above her is exactly the mistake to avoid. Summary lines are
 * skipped so a total is never counted on top of the rows that make it up,
 * and only the last amount on a row is taken, because a row that reads
 * "1 75.00 75.00" is a quantity, a rate and the actual charge.
 */
export function costsIn(
  text: string,
  horses: Horse[]
): { id: string; amount: number; line: string }[] {
  const out: { id: string; amount: number; line: string }[] = [];
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || SUMMARY_LINE.test(line)) continue;

    const amounts = moneyIn(line);
    if (!amounts.length) continue;

    // One horse named twice on a row is still one horse; two different
    // horses on a row is a split we have no basis to make.
    const named = [...new Set(horsesIn(line, horses).map((h) => h.id))];
    if (named.length !== 1) continue;

    out.push({
      id: named[0],
      amount: amounts[amounts.length - 1],
      line: line.slice(0, 120),
    });
  }
  return out;
}

/**
 * What one document says each horse cost.
 *
 * A bill for a single horse usually says so on the outside — the office
 * files it as "Crown Three Vet_10.13.2026_6342_$785.50_Sully Vet.pdf" —
 * and that total is more reliable than anything picked out of the table
 * inside, where a line item may not repeat her name. So when the filename
 * names exactly one horse and carries an amount, that wins outright.
 *
 * Otherwise the document is read as a table, which is the whole-barn case:
 * one farrier visit, twenty rows, a different charge against each name.
 */
export function documentSpend(
  name: string,
  text: string,
  horses: Horse[],
  date: string | null
): Record<string, SpendHit[]> {
  const out: Record<string, SpendHit[]> = {};
  const push = (id: string, amount: number, line: string) => {
    (out[id] ??= []).push({ date, from: name, amount, line });
  };

  const base = name.replace(/\.[a-z0-9]+$/i, "");
  const head = text.slice(0, 1200);
  if (NOT_A_BILL.test(base) || NOT_A_BILL.test(head)) return out;

  const onCover = horsesIn(base, horses);
  const coverIds = [...new Set(onCover.map((h) => h.id))];
  const coverAmounts = moneyIn(base);

  // The office's own filing — one horse and a total on the cover — is
  // itself the statement that this is her bill for that amount.
  if (coverIds.length === 1 && coverAmounts.length) {
    push(coverIds[0], coverAmounts[coverAmounts.length - 1], base);
    return out;
  }

  if (!LOOKS_LIKE_A_BILL.test(base) && !LOOKS_LIKE_A_BILL.test(head)) return out;

  const rows = costsIn(text, horses);
  if (!rows.length) return out;

  // If the cover states a total, the rows have to add up to roughly it.
  // When they do not, the table was misread and a wrong number on Philip's
  // phone is worse than no number.
  const stated = coverAmounts.length
    ? coverAmounts[coverAmounts.length - 1]
    : null;
  const summed = rows.reduce((n, r) => n + r.amount, 0);
  if (stated !== null && Math.abs(summed - stated) > Math.max(1, stated * 0.02)) {
    return {};
  }

  for (const r of rows) push(r.id, r.amount, r.line);
  return out;
}

/* ---------- reading what has been committed ---------- */

/**
 * Where a contract stops stating what was bought and starts listing what
 * could have been bought instead. A stallion contract quotes the fees for
 * every breeding method it does not cover, and counting those would treble
 * the commitment.
 */
const ALTERNATIVES = /\bAlternative Breeding Methods\b/i;

/** The fees a signed stallion contract actually commits to. */
const FEE_LINE =
  /\b(booking|stallion|chute|administration)\s+fee:?\s*\$\s?([\d,]+(?:\.\d{2})?)/gi;

/**
 * What a signed breeding contract commits, and for which mare.
 *
 * A contract is not an invoice: nothing has been billed, but the money is
 * spoken for, and on a mare being bred it is the largest number attached
 * to her. The contract names the stallion rather than the mare — Solo
 * Select's says "Mare: To be designated" — so she is found the other way
 * round, through the stallion named in her own breeding line.
 */
export function contractCommitment(
  text: string,
  horses: Horse[]
): { id: string; amount: number; label: string } | null {
  // Matched from the mare's side rather than the contract's. Pulling the
  // stallion out of the contract means depending on how that vendor lays
  // its header out; every mare being bred already names her stallion in
  // her own breeding line, and the contract either mentions him or does
  // not. One less format to be wrong about.
  const hay = text.toLowerCase();
  let mare: Horse | undefined;
  let stallion = "";
  for (const h of horses) {
    const named = h.breeding?.match(/\bbred to\s+(.+?)(?:\s*\(|\s+at\s+|[,.]|$)/i)?.[1]?.trim();
    if (!named || named.length < 3) continue;
    if (hay.includes(named.toLowerCase())) {
      mare = h;
      stallion = named;
      break;
    }
  }
  if (!mare) return null;

  const head = text.split(ALTERNATIVES)[0];
  const fees = new Map<string, number>();
  for (const m of head.matchAll(FEE_LINE)) {
    const kind = m[1].toLowerCase();
    // First statement of each fee wins; contracts restate them in prose.
    if (!fees.has(kind)) fees.set(kind, toAmount(m[2]));
  }
  const amount = [...fees.values()].reduce((n, v) => n + v, 0);
  if (!amount) return null;

  const parts = [...fees.entries()].map(([k, v]) => `${k} $${v.toLocaleString()}`);
  return { id: mare.id, amount, label: `${stallion} — ${parts.join(", ")}` };
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
  docs: { doc: Doc; ids: string[] }[],
  contracts: Doc[] = []
): Promise<Scan> {
  const items: Scan["items"] = {};
  const spend: Scan["spend"] = {};

  const seen = new Set<string>();
  const pdfs = docs
    .filter((d) => !d.doc.folder && /\.pdf$/i.test(d.doc.name) && d.doc.id)
    // The same invoice can reach here twice — once through a horse's folder
    // and once through Payables. Counting it twice would overstate what she
    // cost, so each file is read once.
    .filter((d) => !seen.has(d.doc.id) && seen.add(d.doc.id))
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

      // What it cost. The invoice date is whatever the care lines agreed on,
      // falling back to a date in the filename.
      const on = found[0]?.date ?? datesIn(entry.doc.name)[0]?.date ?? null;
      const billed = documentSpend(
        entry.doc.name,
        text,
        scoped.length ? scoped : horses,
        on
      );
      for (const [id, hits] of Object.entries(billed)) {
        if (entry.ids.length && !entry.ids.includes(id)) continue;
        (spend[id] ??= []).push(...hits);
      }
    })
  );

  // Signed contracts, read separately. They are full of prices that are
  // not charges — shipping rates, late fees, the methods not purchased —
  // so they never go through the invoice reader.
  const committed: Scan["committed"] = {};
  await Promise.all(
    contracts
      .filter((d) => !d.folder && /\.pdf$/i.test(d.name) && d.id)
      .slice(0, 40)
      .map(async (d) => {
        let text = "";
        try {
          text = await pdfText(d.id);
        } catch {
          return;
        }
        const c = contractCommitment(text, horses);
        if (!c) return;
        (committed[c.id] ??= []).push({
          date: datesIn(d.name)[0]?.date ?? null,
          from: d.name,
          amount: c.amount,
          line: c.label,
        });
      })
  );

  const scan: Scan = {
    at: new Date().toISOString(),
    v: SCAN_VERSION,
    items,
    spend,
    committed,
  };
  try {
    await writeScan(scan);
  } catch {
    /* the scan still returns; it just will not be reused */
  }
  return scan;
}
