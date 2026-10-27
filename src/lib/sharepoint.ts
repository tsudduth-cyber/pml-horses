import { graph, graphAll } from "./graph";
import type { Horse, Doc, Status, CareKey, DerivedCare } from "./types";
import { CARE_TASKS, EMPTY_HORSE } from "./types";
import { readScan, scanIsStale } from "./scan";
import { parseFolder, matchRow } from "./roster";

const HOST = process.env.SP_HOSTNAME!;
const DRIVE = process.env.SP_DRIVE_ID!;
const LIST_NAME = process.env.SP_LIST_NAME ?? "PML Horses";
const ROOT = process.env.SP_HORSES_PATH!;
/** Whole-barn paperwork lives here — one farrier invoice covers everyone. */
const BARN_FOLDER = process.env.SP_BARN_FOLDER ?? "_Barn";
const SOLD_FOLDER = process.env.SP_SOLD_FOLDER ?? "x_Sold Horses";
/**
 * Folders outside /Horses/ that still hold horse paperwork. Payables is
 * filed by month for the books and AQHA by subject — correct for the people
 * who keep them, and not something to reorganise for an app. The app reads
 * them instead and hangs each file on the horse its name mentions.
 */
const SHARED_FOLDERS = (process.env.SP_SHARED_FOLDERS ?? "Payables,AQHA")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);
/** The PML folder that /Horses/ sits inside. Resolved per call, not at
 *  module load — the environment is not there when the build evaluates
 *  this file. */
const pmlRoot = () => ROOT.replace(/\/[^/]+$/, "");

/** Set by getHorses so the page can tell the app whether to rescan. */
let staleScan = false;
let lastDocIndex: { doc: Doc; ids: string[] }[] = [];
export const scanState = () => ({ stale: staleScan, docs: lastDocIndex });

/**
 * Resolve the site and list by name so nobody has to go hunting for GUIDs.
 * Both are stable, so they're looked up once per server process.
 */
let ids: Promise<{ site: string; list: string }> | null = null;

function resolveIds() {
  if (ids) return ids;
  ids = (async () => {
    const site: any = await graph(`/sites/${HOST}`);
    const lists = await graphAll<any>(
      `/sites/${site.id}/lists?$select=id,name,displayName`
    );
    const match = lists.find(
      (l) =>
        l.displayName?.toLowerCase() === LIST_NAME.toLowerCase() ||
        l.name?.toLowerCase() === LIST_NAME.toLowerCase()
    );
    if (!match) {
      throw new Error(
        `LIST_NOT_FOUND: no SharePoint list named "${LIST_NAME}" on ${HOST}`
      );
    }
    return { site: site.id as string, list: match.id as string };
  })().catch((e) => {
    ids = null; // let the next request try again
    throw e;
  });
  return ids;
}

/** Map the list's Status choice text onto the app's status codes. */
const STATUS_MAP: Record<string, Status> = {
  "work in progress": "wip",
  "registered to pml": "done",
  "recip mare": "recip",
  embryo: "exp",
  "not registering": "noreg",
  sold: "off",
};

/**
 * Notes and Missing are one item per line. They split on real newlines — what
 * you get typing into SharePoint — and also on " | ", which is what the seed
 * CSV uses, because grid-view paste treats a newline as a new row.
 */
function lines(v: unknown): string[] {
  if (typeof v !== "string") return [];
  return v
    .split(/\r?\n|\s\|\s/)
    .map((s) => s.replace(/^[-•*]\s*/, "").trim())
    .filter(Boolean);
}

function dateOnly(v: unknown): string | null {
  if (typeof v !== "string" || !v) return null;
  return v.slice(0, 10); // SharePoint hands back ISO; the time is noise here.
}

function num(v: unknown): number | null {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function str(v: unknown): string | null {
  if (typeof v !== "string") return null;
  const s = v.trim();
  return s ? s : null;
}

/** Read every horse record out of the SharePoint list. */
async function readList(): Promise<Horse[]> {
  const { site, list } = await resolveIds();
  const items = await graphAll<any>(
    `/sites/${site}/lists/${list}/items?expand=fields&$top=200`
  );

  return items.map((it) => {
    const f = it.fields ?? {};
    const reg = str(f.Title);
    const barn = str(f.BarnName);
    const pending = String(f.NamePending ?? "").toLowerCase() === "true" ||
      f.NamePending === true;

    const care: Partial<Record<CareKey, string>> = {};
    for (const [k, col] of [
      ["farrier", "Farrier"],
      ["worm", "Deworm"],
      ["vax", "Vax"],
      ["coggins", "Coggins"],
      ["dental", "Dental"],
    ] as const) {
      const d = dateOnly(f[col]);
      if (d) care[k] = d;
    }

    const folderName = str(f.FolderName);

    return {
      id: String(it.id),
      reg,
      barn,
      pending,
      idline: str(f.IdLine),
      aqha: str(f.AQHA),
      sex: str(f.Sex),
      color: str(f.Color),
      dob: dateOnly(f.Foaled),
      bought: dateOnly(f.Bought),
      boughtNote: str(f.BoughtNote),
      price: num(f.Price),
      salePrice: num(f.SalePrice),
      soldOn: dateOnly(f.SoldOn),
      status: STATUS_MAP[String(f.Status ?? "").toLowerCase()] ?? "done",
      label: str(f.StatusNote),
      pedigree: str(f.Pedigree),
      breeding: str(f.BreedingTo),
      notes: lines(f.Notes),
      missing: lines(f.Missing),
      care,
      derived: {},
      folderName,
      folderUrl: null,
      docs: [],
      photos: [],
      sort: num(f.Sort) ?? 999,
    } satisfies Horse;
  });
}

function toDoc(c: any): Doc {
  return {
    id: String(c.id ?? ""),
    name: c.name,
    url: c.webUrl,
    folder: !!c.folder,
    modified: c.lastModifiedDateTime ?? "",
  };
}

/**
 * Read one horse's folder. A missing folder is normal — a recip mare or an
 * embryo may not have one — so it comes back empty rather than throwing.
 */
async function readFolder(folderName: string) {
  const path = encodeURIComponent(`${ROOT}/${folderName}`)
    .replace(/%2F/g, "/");
  try {
    const res: any = await graph(
      `/drives/${DRIVE}/root:/${path}:/children?$top=200`
    );
    const children: any[] = res.value ?? [];

    const photosFolder = children.find(
      (c) => c.folder && /^photos$/i.test(c.name)
    );

    let photos: Doc[] = [];
    if (photosFolder) {
      const p: any = await graph(
        `/drives/${DRIVE}/items/${photosFolder.id}/children?$top=200`
      );
      photos = (p.value ?? [])
        .filter((c: any) => !c.folder && /\.(jpe?g|png|heic|webp)$/i.test(c.name))
        .map(toDoc)
        .sort((a: Doc, b: Doc) => a.name.localeCompare(b.name));
    }

    const docs = children
      .filter((c) => c !== photosFolder)
      .map(toDoc)
      .sort((a, b) => Number(b.folder) - Number(a.folder) || a.name.localeCompare(b.name));

    return { docs, photos };
  } catch {
    return { docs: [] as Doc[], photos: [] as Doc[] };
  }
}

/**
 * Read a folder and everything under it.
 *
 * Payables is filed by month and AQHA by subject, so the paperwork for a
 * horse sits a level or two down. Depth is capped because this walks on
 * every page load and nothing here is filed deeper than that.
 */
async function readTree(path: string, depth = 2): Promise<Doc[]> {
  const enc = encodeURIComponent(path).replace(/%2F/g, "/");
  let children: any[] = [];
  try {
    const res: any = await graph(`/drives/${DRIVE}/root:/${enc}:/children?$top=200`);
    children = res.value ?? [];
  } catch {
    return []; // a folder that is not there is not an error
  }

  const files = children.filter((c) => !c.folder).map(toDoc);
  if (depth <= 0) return files;

  const deeper = await Promise.all(
    children
      .filter((c) => c.folder)
      .map((c) => readTree(`${path}/${c.name}`, depth - 1))
  );
  return files.concat(...deeper);
}

/**
 * Pull a date out of a filename. The filing standard is m.d.yy — "Sully
 * Coggins 10.13.26.pdf" — and that is what this looks for first. It also
 * takes m-d-yy and an ISO date, because people type what they type.
 *
 * Returns null rather than guessing. A wrong care date is worse than none:
 * a wrong Coggins reads as covered when the horse is not.
 */
export function dateInName(name: string): string | null {
  const iso = name.match(/\b(20\d{2})-(\d{1,2})-(\d{1,2})\b/);
  if (iso) {
    const [, y, m, d] = iso;
    return ymd(+y, +m, +d);
  }

  const us = name.match(/\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2}|\d{4})\b/);
  if (us) {
    const [, mm, dd, yy] = us;
    const year = yy.length === 2 ? 2000 + +yy : +yy;
    return ymd(year, +mm, +dd);
  }
  return null;
}

function ymd(y: number, m: number, d: number): string | null {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  // Reject a date that rolled over — 2.30.26 is a typo, not Mar 2.
  if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  // Nothing in the future, and nothing absurdly old.
  const now = Date.now();
  if (dt.getTime() > now + 864e5 || y < 2000) return null;
  return dt.toISOString().slice(0, 10);
}

/**
 * Work out care dates from the paperwork on file.
 *
 * Kristina files a vet or farrier bill with the date in the name; this reads
 * it so nobody has to key the same date twice. Only the newest file for a
 * task counts. Barn-wide paperwork — one farrier visit, one invoice — lives
 * in the _Barn folder and applies to every horse.
 */
function deriveCare(
  docs: Doc[],
  barn: boolean
): Partial<Record<CareKey, DerivedCare>> {
  const out: Partial<Record<CareKey, DerivedCare>> = {};
  for (const e of careEntries(docs)) {
    const held = out[e.key];
    if (!held || e.date > held.date) {
      out[e.key] = { date: e.date, from: e.from, barn };
    }
  }
  return out;
}

/** One piece of care paperwork: what it was, when, and the file it came from. */
interface CareEntry {
  key: CareKey;
  date: string;
  from: string;
  /** The filename without its extension, for matching horse names against. */
  base: string;
}

export function careEntries(docs: Doc[]): CareEntry[] {
  const out: CareEntry[] = [];
  for (const d of docs) {
    if (d.folder) continue;
    // A picture is never proof that work was done. "Shots 10.13.26.jpg" is
    // a photo of a horse, not a vaccination record.
    if (/\.(jpe?g|png|heic|webp|gif|mov|mp4)$/i.test(d.name)) continue;
    const base = d.name.replace(/\.[a-z0-9]+$/i, "");
    const date = dateInName(base);
    if (!date) continue;
    for (const t of CARE_TASKS) {
      if (t.match.test(base)) out.push({ key: t.key, date, from: d.name, base });
    }
  }
  return out;
}

const rxEscape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/**
 * Which horses a piece of barn paperwork is about.
 *
 * A vet comes out and does six head on one invoice, so the filename can name
 * them — "Vet Vax 9.2.26 - Sully, Roanie, Darlin". Naming nobody means the
 * whole barn, which is the farrier case. Returns null for "everyone".
 */
export function namedHorses(base: string, horses: Horse[]): Set<string> | null {
  const hit = new Set<string>();
  for (const h of horses) {
    for (const name of [h.barn, h.reg]) {
      // Two characters is enough — JP is a real barn name here, and missing
      // it would spread his invoice across the whole place.
      if (!name || name.length < 2) continue;
      if (new RegExp(`(^|[^a-z0-9])${rxEscape(name)}([^a-z0-9]|$)`, "i").test(base)) {
        hit.add(h.id);
        break;
      }
    }
  }
  return hit.size ? hit : null;
}

/** Build the web URL for a horse's folder so "Open folder" goes somewhere real. */
function folderWebUrl(folderName: string): string {
  const base = process.env.SP_WEB_BASE!.replace(/\/$/, "");
  return `${base}/${ROOT}/${folderName}`
    .split("/")
    .map((seg, i) => (i < 3 ? seg : encodeURIComponent(seg)))
    .join("/");
}

/**
 * Everything the app needs, in one pass: the list for the facts, each folder
 * for the documents and photos. Folders are read in parallel.
 */

/** Folder names under /Horses/, including the sold ones a level down. */
async function horseFolders(): Promise<string[]> {
  const list = async (path: string) => {
    try {
      const enc = encodeURIComponent(path).replace(/%2F/g, "/");
      const res: any = await graph(`/drives/${DRIVE}/root:/${enc}:/children?$top=200`);
      return (res.value ?? []).filter((c: any) => c.folder).map((c: any) => c.name as string);
    } catch {
      return [] as string[];
    }
  };
  const top = await list(ROOT);
  const sold = top.includes(SOLD_FOLDER) ? await list(`${ROOT}/${SOLD_FOLDER}`) : [];
  return [
    // Folders starting with _ are the app's own and shared material.
    ...top.filter((n: string) => !n.startsWith("_") && n !== SOLD_FOLDER),
    ...sold.map((n: string) => `${SOLD_FOLDER}/${n}`),
  ];
}

export async function getHorses(): Promise<Horse[]> {
  const [rows, folders, barnDocs, sharedDocs] = await Promise.all([
    readList(),
    horseFolders(),
    // One farrier visit does the whole barn and produces one invoice, so
    // whole-barn paperwork is filed once here rather than copied 22 times.
    readFolder(BARN_FOLDER).then((f) => f.docs),
    Promise.all(
      SHARED_FOLDERS.map((f) => readTree(`${pmlRoot()}/${f}`))
    ).then((r) => r.flat()),
  ]);

  /**
   * The roster comes from the folders, not from the list.
   *
   * A folder is a horse and its name is her name, which is the convention
   * the office already files by. The list supplies what a folder cannot —
   * price, status, care dates, notes — attached by name so that renaming a
   * folder renames the horse instead of orphaning her record. A folder with
   * no record still shows: that is how a new horse arrives.
   */
  const claimed = new Set<Horse>();
  const horses: Horse[] = folders.map((folder) => {
    const p = parseFolder(folder);
    const row = matchRow(folder, p, rows, claimed);
    if (row) claimed.add(row);

    const base: Horse = row
      ? { ...row }
      : {
          ...EMPTY_HORSE,
          id: `folder:${folder}`,
          status: folder.startsWith(`${SOLD_FOLDER}/`) ? "off" : "wip",
        };

    return {
      ...base,
      // The folder wins on identity. That is the whole point.
      reg: p.reg ?? base.reg,
      barn: p.barn ?? base.barn,
      pending: p.pending || (!p.reg && !p.barn && base.pending),
      idline: p.idline ?? base.idline,
      folderName: folder,
      status: folder.startsWith(`${SOLD_FOLDER}/`) ? "off" : base.status,
    };
  });

  // Records with no folder of their own — embryos, a recip mare — still
  // belong on the list. They are horses, they just have no paperwork yet.
  horses.push(...rows.filter((r) => !claimed.has(r)).map((r) => ({ ...r, folderName: null })));

  // A shared document belongs to whichever horses its filename mentions.
  // One that names nobody belongs to nobody — unlike barn paperwork, a
  // credit card statement is not everyone's.
  const shared = sharedDocs
    .map((d) => ({ doc: d, only: namedHorses(d.name, horses) }))
    .filter((e) => e.only);

  // Work out who each barn document is about once, not once per horse.
  const barn = careEntries(barnDocs).map((e) => ({
    ...e,
    only: namedHorses(e.base, horses),
  }));

  await Promise.all(
    horses.map(async (h) => {
      // Barn paperwork first, so the horse's own file beats it.
      const from: Partial<Record<CareKey, DerivedCare>> = {};
      for (const e of barn) {
        if (e.only && !e.only.has(h.id)) continue;
        const held = from[e.key];
        if (!held || e.date > held.date) {
          from[e.key] = { date: e.date, from: e.from, barn: true };
        }
      }

      if (!h.folderName) {
        // No folder of her own, but shared paperwork can still name her.
        const mine = shared.filter((e) => e.only!.has(h.id)).map((e) => e.doc);
        h.docs = mine;
        h.derived = { ...from, ...deriveCare(mine, false) };
        return;
      }
      const { docs, photos } = await readFolder(h.folderName);
      const mine = shared.filter((e) => e.only!.has(h.id)).map((e) => e.doc);
      h.docs = [...docs, ...mine].sort(
        (a, b) =>
          Number(b.folder) - Number(a.folder) || a.name.localeCompare(b.name)
      );
      h.photos = photos;
      h.folderUrl = folderWebUrl(h.folderName);
      h.derived = { ...from, ...deriveCare([...docs, ...mine], false) };
    })
  );

  // What the scanner needs: every document, and whose it is.
  lastDocIndex = [
    ...horses.flatMap((h) =>
      h.docs.filter((d) => !d.folder).map((doc) => ({ doc, ids: [h.id] }))
    ),
    // Barn paperwork is not on any one horse's page, so it would otherwise
    // never be read. No ids means "whoever the document names inside".
    ...barnDocs.filter((d) => !d.folder).map((doc) => ({ doc, ids: [] as string[] })),
  ];

  // Care dates read out of the documents themselves beat ones guessed from a
  // filename — "Sully Vet.pdf" says nothing, the invoice inside says Coggins.
  const scan = await readScan();
  if (scan) {
    for (const h of horses) {
      const hits = scan.items[h.id];
      if (!hits) continue;
      for (const [k, hit] of Object.entries(hits)) {
        h.derived[k as CareKey] = {
          date: hit.date,
          from: hit.from,
          barn: false,
        };
      }
    }
  }
  staleScan = scanIsStale(scan);

  return horses.sort(
    (a, b) =>
      a.sort - b.sort ||
      (a.barn ?? a.reg ?? "").localeCompare(b.barn ?? b.reg ?? "")
  );
}

/**
 * The list item id to write to.
 *
 * A horse found as a folder has no record yet — that is normal, it is how
 * a new horse arrives. The first time anyone writes to her, give her one,
 * pointed at her folder. Nobody has to create a record by hand.
 */
async function resolveItemId(id: string): Promise<string> {
  if (!id.startsWith("folder:")) return id;
  const folder = id.slice("folder:".length);
  const { site, list } = await resolveIds();

  const existing = await graphAll<any>(
    `/sites/${site}/lists/${list}/items?expand=fields&$top=200`
  );
  const hit = existing.find(
    (it) =>
      String(it.fields?.FolderName ?? "").toLowerCase() === folder.toLowerCase()
  );
  if (hit) return String(hit.id);

  const made: any = await graph(`/sites/${site}/lists/${list}/items`, {
    method: "POST",
    body: JSON.stringify({ fields: { FolderName: folder, Sort: 900 } }),
  });
  return String(made.id);
}

const CARE_COLUMN: Record<CareKey, string> = {
  farrier: "Farrier",
  worm: "Deworm",
  vax: "Vax",
  coggins: "Coggins",
  dental: "Dental",
};

/**
 * Write one care date back to the list. This is the only field the app
 * writes — registration and paperwork stay with whoever does the filing.
 */
export async function saveCareDate(
  itemId: string,
  key: CareKey,
  date: string | null
) {
  const column = CARE_COLUMN[key];
  if (!column) throw new Error("BAD_FIELD");
  if (date !== null && !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    throw new Error("BAD_DATE");
  }

  const row = await resolveItemId(itemId);
  const { site, list } = await resolveIds();
  await graph(`/sites/${site}/lists/${list}/items/${row}/fields`, {
    method: "PATCH",
    body: JSON.stringify({ [column]: date }),
  });
}

/** The text fields the app may rewrite. Everything else stays read-only. */
const TEXT_COLUMN = { notes: "Notes", missing: "Missing" } as const;
export type TextKey = keyof typeof TEXT_COLUMN;

/**
 * Rewrite a horse's notes.
 *
 * The office kept these in a Word document because the app gave them
 * nowhere else to write, and the two drifted apart over ten days. This is
 * the somewhere else. Lines are stored one per line, which is what the
 * reader already splits on.
 */
export async function saveLines(itemId: string, key: TextKey, lines: string[]) {
  const column = TEXT_COLUMN[key];
  if (!column) throw new Error("BAD_FIELD");

  const clean = lines
    .map((l) => l.replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .slice(0, 40);
  if (clean.some((l) => l.length > 600)) throw new Error("TOO_LONG");

  const row = await resolveItemId(itemId);
  const { site, list } = await resolveIds();
  await graph(`/sites/${site}/lists/${list}/items/${row}/fields`, {
    method: "PATCH",
    body: JSON.stringify({ [column]: clean.join("\n") }),
  });
  return clean;
}

/**
 * The record fields the app may rewrite, and the column each one lands in.
 *
 * Everything here used to be seeded once and then frozen — a horse could be
 * renamed in SharePoint, given a folder and photographed, and the app would
 * go on showing the old record because nobody could reach it. This is how
 * it gets reached.
 */
const RECORD_COLUMN = {
  barn: "BarnName",
  reg: "Title",
  aqha: "AQHA",
  sex: "Sex",
  color: "Color",
  dob: "Foaled",
  bought: "Bought",
  folderName: "FolderName",
} as const;
export type RecordKey = keyof typeof RECORD_COLUMN;

const SEXES = ["Stallion", "Mare", "Gelding", "Filly", "Colt"];
const DATE_KEYS: RecordKey[] = ["dob", "bought"];

export async function saveRecord(
  itemId: string,
  patch: Partial<Record<RecordKey, string | null>>
) {
  const fields: Record<string, unknown> = {};

  for (const [k, v] of Object.entries(patch) as [RecordKey, string | null][]) {
    const column = RECORD_COLUMN[k];
    if (!column) throw new Error("BAD_FIELD");
    const val = typeof v === "string" ? v.trim() : v;

    if (!val) {
      fields[column] = null;
      continue;
    }
    if (DATE_KEYS.includes(k) && !/^\d{4}-\d{2}-\d{2}$/.test(val)) {
      throw new Error("BAD_DATE");
    }
    if (k === "sex" && !SEXES.includes(val)) throw new Error("BAD_SEX");
    if (val.length > 255) throw new Error("TOO_LONG");
    // Dates go in at noon UTC so no timezone can roll them back a day.
    fields[column] = DATE_KEYS.includes(k) ? `${val}T12:00:00Z` : val;
  }

  // A horse with a registered name is no longer awaiting one.
  if ("reg" in patch) fields.NamePending = !patch.reg;

  const row = await resolveItemId(itemId);
  const { site, list } = await resolveIds();

  // Pointing a record at a folder another record already holds is a merge,
  // not a rename. Two records for one horse is exactly the mess this is
  // meant to clear up, and the stray is usually the one carrying the note
  // nobody wants to lose — so its content moves over before it goes.
  const moveTo =
    typeof patch.folderName === "string" ? patch.folderName.trim() : "";
  if (moveTo) {
    const held = await rowHoldingFolder(moveTo, row);
    if (held) {
      await mergeRows(held, row, fields);
      return;
    }
    // Moving a record onto a folder renames the horse — that is the point
    // of folders driving the roster. A barn name the new folder doesn't
    // mention is left over from the old one ("New recip" on the Appaloosa
    // Recip folder), and keeping it would leave her listed under the wrong
    // name after the very save meant to fix that. A nickname the folder
    // does carry stays.
    if (!("barn" in patch)) {
      const current = await graph<any>(
        `/sites/${site}/lists/${list}/items/${row}?expand=fields`
      ).catch(() => null);
      const was = String(current?.fields?.BarnName ?? "").trim();
      if (was && !moveTo.toLowerCase().includes(was.toLowerCase())) {
        fields.BarnName = null;
      }
    }
  }

  await graph(`/sites/${site}/lists/${list}/items/${row}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fields),
  });
}

/** The list item already pointing at `folder`, if it isn't `exceptId`. */
async function rowHoldingFolder(folder: string, exceptId: string) {
  const { site, list } = await resolveIds();
  const items = await graphAll<any>(
    `/sites/${site}/lists/${list}/items?expand=fields&$top=200`
  );
  return (
    items.find(
      (it) =>
        String(it.id) !== String(exceptId) &&
        String(it.fields?.FolderName ?? "").toLowerCase() ===
          folder.toLowerCase()
    ) ?? null
  );
}

/**
 * Columns worth carrying when two records turn out to be one horse.
 *
 * Name and barn name are deliberately absent. The target is the record
 * holding the folder, and the folder is what names a horse — carrying a
 * stale "New recip" over the top of "Appaloosa Recip" would undo the very
 * thing the merge is for. The source contributes history, not identity.
 */
const MERGE_COLUMNS = [
  "IdLine", "AQHA", "Sex", "Color", "Foaled",
  "Bought", "BoughtNote", "Price", "SalePrice", "SoldOn", "Status",
  "StatusNote", "Pedigree", "BreedingTo", "Farrier", "Deworm", "Vax",
  "Coggins", "Dental",
];

const empty = (v: unknown) =>
  v === null || v === undefined || String(v).trim() === "";

/**
 * Fold `sourceId` into `targetId` and delete the source.
 *
 * The target keeps everything it already has — it is the one with the
 * folder, so it is the one the office is maintaining. The source fills in
 * the blanks, and notes are unioned rather than chosen between, because a
 * line like "believed to carry Embryo 2" exists in exactly one place.
 *
 * The copy is confirmed before the delete. If the delete fails the stray
 * record simply shows up again, which is visible and fixable; the reverse
 * order would lose it.
 */
async function mergeRows(
  target: any,
  sourceId: string,
  pending: Record<string, unknown>
) {
  const { site, list } = await resolveIds();
  const src: any = await graph(
    `/sites/${site}/lists/${list}/items/${sourceId}?expand=fields`
  );
  const from = { ...(src.fields ?? {}), ...pending };
  const have = target.fields ?? {};

  const patch: Record<string, unknown> = {};
  for (const col of MERGE_COLUMNS) {
    if (empty(have[col]) && !empty(from[col])) patch[col] = from[col];
  }

  for (const col of ["Notes", "Missing"]) {
    const merged = [...lines(have[col]), ...lines(from[col])];
    const seen = new Set<string>();
    const kept = merged.filter((l) => {
      const k = l.trim().toLowerCase();
      if (!k || seen.has(k)) return false;
      seen.add(k);
      return true;
    });
    if (kept.length > lines(have[col]).length) patch[col] = kept.join("\n");
  }

  if (Object.keys(patch).length) {
    await graph(`/sites/${site}/lists/${list}/items/${target.id}/fields`, {
      method: "PATCH",
      body: JSON.stringify(patch),
    });
  }

  await graph(`/sites/${site}/lists/${list}/items/${sourceId}`, {
    method: "DELETE",
  });
}

/**
 * Every horse folder under /Horses/, for the folder picker.
 *
 * This used to list only the folders no record pointed at, which made sense
 * when the list was the roster. Now the folders are the roster, so every
 * folder is claimed by its own horse and that list was always empty — there
 * was no way to put a stray record onto the folder it belongs to. Offering
 * all of them is what makes the merge reachable.
 */
export async function allFolders(): Promise<string[]> {
  return (await horseFolders()).sort();
}
