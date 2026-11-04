import type { Horse } from "./types";

/**
 * Reading the office's "Horse Info & Registration Status" document.
 *
 * AQHA numbers, foaling dates, purchase dates and prices are not on any
 * machine-readable paperwork: the older registration certificates are
 * image scans, and the bills of sale mostly are too. They exist in one
 * place, which is the document Kristina keeps, as a numbered list:
 *
 *   12. Chasin Checks: 6424932
 *       a. Printed digital certificate
 *       b. Stallion
 *       c. Date of Sale: 8/17/26?
 *       d. DOB: 2/20/25
 *       e. Price: $12,500
 *
 * The app reads that whenever it changes and fills in whatever it is
 * missing. It never overwrites something the app already holds, so a
 * value entered in the app wins, and it never writes a value it could
 * not read with confidence. The text comes out of Word with spaces
 * inside numbers ("8/1 7 /26", "$30 ,000"), which is handled.
 */

export interface RecordEntry {
  /** The registered name as written, e.g. "Blazin Hot Metal". */
  name: string;
  /** The whole entry, for matching a horse the header does not name. */
  raw: string;
  /** The barn name in parentheses, when there is one. */
  barn: string | null;
  aqha: string | null;
  sex: string | null;
  color: string | null;
  /** ISO dates. */
  dob: string | null;
  bought: string | null;
  price: number | null;
}

const MONTHS = "jan feb mar apr may jun jul aug sep oct nov dec".split(" ");

/** Collapse the stray spaces Word's text layer leaves inside tokens. */
function tidy(s: string) {
  return s
    .replace(/(\d)\s+(?=[\d/,.])/g, "$1")
    .replace(/([/,.$])\s+(?=\d)/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

function iso(y: number, m: number, d: number): string | null {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  if (y < 100) y += y < 50 ? 2000 : 1900;
  return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

/** The first date in a string, numeric or written out. */
function dateIn(s: string): string | null {
  const t = tidy(s);
  const num = t.match(/\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b/);
  const long = t.match(/\b([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s*(\d{4})\b/);
  const pick = (m: RegExpMatchArray | null, kind: "num" | "long") => {
    if (!m) return null;
    if (kind === "num") return iso(+m[3], +m[1], +m[2]);
    const mi = MONTHS.indexOf(m[1].slice(0, 3).toLowerCase());
    return mi < 0 ? null : iso(+m[3], mi + 1, +m[2]);
  };
  // Whichever comes first in the text.
  if (num && long) {
    return num.index! <= long.index! ? pick(num, "num") : pick(long, "long");
  }
  return pick(num, "num") ?? pick(long, "long");
}

const SEXES: [RegExp, string][] = [
  [/\bgelded\b|\bgelding\b/i, "Gelding"],
  [/\bstallion\b/i, "Stallion"],
  [/\bfilly\b/i, "Filly"],
  [/\bcolt\b/i, "Colt"],
  [/\bmare\b/i, "Mare"],
];

const COLORS =
  /\b(sorrel|chestnut|bay|black|brown|grey|gray|roan|red roan|bay roan|blue roan|palomino|buckskin|dun|grulla|grullo)\b/i;

/** One entry of the list: the header line and its lettered sub-items. */
function parseEntry(block: string): RecordEntry | null {
  // Header runs up to the first lettered sub-item.
  const split = block.search(/\s[a-f]\.\s/);
  const head = tidy(split >= 0 ? block.slice(0, split) : block);
  const body = split >= 0 ? block.slice(split) : "";

  // "Blazin Hot Metal ( Sully ) : 6493604 . Mailed check..."
  const m = head.match(/^(.+?)(?:\s*\(([^)]+)\))?\s*:/);
  if (!m) return null;
  const name = tidy(m[1]);
  let barn: string | null = m[2] ? tidy(m[2]) : null;
  // "(Owner didn't give her a barn name)" is a remark, not a name.
  if (barn && barn.split(" ").length > 3) barn = null;
  if (!name || name.length < 2) return null;

  const aqha = head.match(/\b(\d{7})\b/)?.[1] ?? null;

  const field = (label: RegExp) => {
    const r = body.match(label);
    return r ? r[1] : null;
  };
  const dobRaw = field(/\bDOB:\s*([^\n]*?)(?=\s+[a-f]\.\s|\s*$)/i);
  const soldRaw = field(/\bDate of Sale:\s*([^\n]*?)(?=\s+[a-f]\.\s|\s*$)/i);
  const priceRaw = field(/\bPrice\s*:\s*([^\n]*?)(?=\s+[a-f]\.\s|\s*$)/i);
  const colorRaw = field(/\bColor:\s*([A-Za-z ]+)/i);

  let price: number | null = null;
  if (priceRaw) {
    const t = tidy(priceRaw).match(/\$\s*([\d,]+(?:\.\d{2})?)/);
    if (t) price = Number(t[1].replace(/,/g, ""));
  }

  // Only a sub-item that IS the sex counts — "b. Stallion", "a. Sorrel
  // Stallion", "a. Gelded: Yes". Prose like "in a mare at PL's ranch"
  // under an embryo must not make the embryo a mare.
  let sex: string | null = null;
  const item = body.match(
    /\s[a-f]\.\s+(?:[A-Za-z]+\s+)?(gelded|gelding|stallion|mare|filly|colt)\b/i
  );
  if (item) {
    for (const [re, label] of SEXES) {
      if (re.test(item[1])) {
        sex = label;
        break;
      }
    }
  }

  let color: string | null = colorRaw ? tidy(colorRaw).split(" ")[0] : null;
  if (!color) {
    // "a. Sorrel Stallion" — the colour rides in front of the sex.
    const c = body.match(COLORS);
    if (c && /\b(stallion|mare|gelding|filly|colt)\b/i.test(body.slice(c.index!, c.index! + 30))) {
      color = c[1];
    }
  }
  if (color) color = color[0].toUpperCase() + color.slice(1).toLowerCase();

  return {
    name,
    raw: tidy(block),
    barn,
    aqha,
    sex,
    color,
    dob: dobRaw ? dateIn(dobRaw) : null,
    bought: soldRaw ? dateIn(soldRaw) : null,
    price: Number.isFinite(price as number) ? price : null,
  };
}

/** Every horse the document describes. */
export function parseHorseInfo(text: string): RecordEntry[] {
  // Entries begin with a number and a period, sub-items with a letter.
  // Split on the numbers, keeping each entry's body with it.
  const parts = text.split(/(?:^|\s)(?=\d{1,2}\.\s+[A-Z])/);
  const out: RecordEntry[] = [];
  for (const p of parts) {
    const body = p.replace(/^\s*\d{1,2}\.\s+/, "");
    if (body === p) continue; // did not start with a number
    const e = parseEntry(body);
    if (e) out.push(e);
  }
  return out;
}

const key = (s: string | null | undefined) =>
  (s ?? "")
    .toLowerCase()
    .replace(/['’]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

/** The same with every space gone, so Word's "E mbryo 1" still reads. */
const squash = (s: string | null | undefined) => key(s).replace(/ /g, "");

/** The horse on the roster an entry describes, or nothing. */
export function matchEntry(e: RecordEntry, horses: Horse[]): Horse | null {
  const n = squash(e.name);
  const b = squash(e.barn);
  const raw = key(e.raw);
  for (const h of horses) {
    const names = [h.reg, h.barn, h.folderName?.split("/").pop(), ...h.aka]
      .map(squash)
      .filter(Boolean);
    if (names.includes(n)) return h;
    if (b && names.includes(b)) return h;
    // A horse with no name yet is on the roster by her parents — "Dual
    // Reyish x High Brow Madonna" — and the entry names them in prose.
    if (h.pending && h.idline) {
      const sire = key(h.idline).split(" x ")[0];
      if (sire && sire.split(" ").length >= 2 && raw.includes(sire)) return h;
    }
    // The folder "Blazin Hot Metal (Sully)" carries both at once.
    const folder = squash(h.folderName?.split("/").pop());
    if (folder && folder.includes(n) && (!b || folder.includes(b))) return h;
    // A horse with no name yet is told apart by her line — "Sue's Baby".
    if (h.idline && squash(h.idline).includes(n)) return h;
  }
  return null;
}

/**
 * The list columns this entry can fill on a horse that lacks them. Only
 * blanks: nothing the app already holds is touched.
 */
export function blanksToFill(
  e: RecordEntry,
  h: Horse
): Record<string, unknown> {
  const f: Record<string, unknown> = {};
  if (e.aqha && !h.aqha) f.AQHA = e.aqha;
  if (e.sex && !h.sex) f.Sex = e.sex;
  if (e.color && !h.color) f.Color = e.color;
  if (e.dob && !h.dob) f.Foaled = `${e.dob}T12:00:00Z`;
  if (e.bought && !h.bought) f.Bought = `${e.bought}T12:00:00Z`;
  if (e.price !== null && h.price == null) f.Price = e.price;
  return f;
}

/* ---------- AQHA registration certificates ---------- */

/** Does this text read as an AQHA certificate of registration? */
export function isCertificate(text: string) {
  return (
    /American Quarter Horse Association/i.test(text) &&
    /REGISTRATION NUMBER/i.test(text)
  );
}

/**
 * The record off a digital AQHA certificate.
 *
 * The first line carries everything in a fixed order — name, foaling
 * date, colour, sex, registration number, then the sire — and the back
 * repeats the name and number under labels. For a horse that arrives
 * from now on this is the source, read from the certificate in her own
 * folder the day it is filed, and nobody types anything. The older
 * certificates are image scans with no text, which is what the Horse
 * Info document covers.
 */
export function certificateEntry(text: string): RecordEntry | null {
  if (!isCertificate(text)) return null;
  const t = text.replace(/\s+/g, " ");

  const aqha = t.match(/REGISTRATION NUMBER\s+(\d{7})\b/i)?.[1] ?? null;
  const name = t.match(/REGISTERED NAME\s+([A-Z][A-Z' ]+?)\s+(?:DATE ISSUED|REGISTRATION)/)?.[1]?.trim() ?? null;
  if (!aqha || !name) return null;

  // "VINTAGE ENTERPRIZE May 28, 2021 Bay GELDING 6132734"
  const head = t.match(
    new RegExp(
      `([A-Za-z]{3,9}\\s+\\d{1,2},\\s*\\d{4})\\s+([A-Za-z ]+?)\\s+(STALLION|MARE|GELDING|FILLY|COLT)\\s+${aqha}\\b`
    )
  );
  const dob = head ? dateIn(head[1]) : null;
  const color = head ? head[2].trim() : null;
  const sexWord = head ? head[3] : t.match(/\b(STALLION|MARE|GELDING|FILLY|COLT)\b/)?.[1] ?? null;
  const sex = sexWord ? sexWord[0] + sexWord.slice(1).toLowerCase() : null;

  const acquired = t.match(/DATE ACQUIRED\s+([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})/i)?.[1];

  return {
    name: name
      .toLowerCase()
      .replace(/\b[a-z]/g, (c) => c.toUpperCase()),
    raw: t,
    barn: null,
    aqha,
    sex,
    color: color ? color[0].toUpperCase() + color.slice(1).toLowerCase() : null,
    dob,
    bought: acquired ? dateIn(acquired) : null,
    price: null,
  };
}
