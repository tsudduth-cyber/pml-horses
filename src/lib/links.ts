import { graph, graphRaw } from "./graph";

const DRIVE = process.env.SP_DRIVE_ID!;
const ROOT = process.env.SP_HORSES_PATH!;

/** Where the app remembers which folder belongs to which record. */
const PATH = () => `${ROOT}/_app/links.json`;

/**
 * Folder id → list row id.
 *
 * Matching a folder to its record by name works right up until the name
 * changes, which is the one thing a horse's folder reliably does. Lizard
 * survived his rename only because "Lizard" stayed in the new folder
 * name; the day Philip names a horse something with no overlap at all —
 * Appaloosa Recip becoming Smart Little Lena — the name tells us nothing
 * and the record orphans.
 *
 * A SharePoint folder keeps its id through a rename. So the link is kept
 * on the id, learned the first time a folder and a record are matched by
 * name, and from then on the folder can be called anything at all.
 */
export interface Link {
  /** The list row this folder belongs to. */
  row: string;
  /** What the folder was called when we last looked. */
  name: string;
}

export type Links = Record<string, Link>;

export async function readLinks(): Promise<Links> {
  try {
    const p = encodeURIComponent(PATH()).replace(/%2F/g, "/");
    const res = await graphRaw(`/drives/${DRIVE}/root:/${p}:/content`);
    const data = JSON.parse(await res.text());
    if (data && typeof data === "object" && !Array.isArray(data)) {
      // Drop anything that is not the shape we expect, so an older or
      // half-written file degrades to "nothing learned" rather than
      // throwing on every page load.
      const out: Links = {};
      for (const [k, v] of Object.entries(data as Record<string, unknown>)) {
        const l = v as Partial<Link>;
        if (l && typeof l.row === "string" && typeof l.name === "string") {
          out[k] = { row: l.row, name: l.name };
        }
      }
      return out;
    }
  } catch {
    /* nothing learned yet */
  }
  return {};
}

/**
 * Save the map, but only when it actually changed — this runs off a page
 * load, and most page loads learn nothing new. Failure is silent: a lost
 * link costs one rename's worth of continuity, a thrown error costs the
 * page.
 */
export async function writeLinks(next: Links, prev: Links) {
  const same =
    Object.keys(next).length === Object.keys(prev).length &&
    Object.entries(next).every(
      ([k, v]) => prev[k]?.row === v.row && prev[k]?.name === v.name
    );
  if (same) return;

  try {
    const p = encodeURIComponent(PATH()).replace(/%2F/g, "/");
    await graph(`/drives/${DRIVE}/root:/${p}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(next, null, 1),
    });
  } catch {
    /* best effort */
  }
}
