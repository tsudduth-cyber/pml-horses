import { graph, graphAll } from "./graph";
import type { Horse, Doc, Status, CareKey } from "./types";

const HOST = process.env.SP_HOSTNAME!;
const DRIVE = process.env.SP_DRIVE_ID!;
const LIST_NAME = process.env.SP_LIST_NAME ?? "PML Horses";
const ROOT = process.env.SP_HORSES_PATH!;

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
export async function getHorses(): Promise<Horse[]> {
  const horses = await readList();

  await Promise.all(
    horses.map(async (h) => {
      if (!h.folderName) return;
      const { docs, photos } = await readFolder(h.folderName);
      h.docs = docs;
      h.photos = photos;
      h.folderUrl = folderWebUrl(h.folderName);
    })
  );

  return horses.sort(
    (a, b) =>
      a.sort - b.sort ||
      (a.barn ?? a.reg ?? "").localeCompare(b.barn ?? b.reg ?? "")
  );
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

  const { site, list } = await resolveIds();
  await graph(`/sites/${site}/lists/${list}/items/${itemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify({ [column]: date }),
  });
}
