import { graph, graphRaw } from "./graph";

const DRIVE = process.env.SP_DRIVE_ID!;
const ROOT = process.env.SP_HORSES_PATH!;

/** Where the app keeps the names a horse used to go by. */
const PATH = () => `${ROOT}/_app/aliases.json`;

/**
 * A horse's former names, so her old paperwork keeps finding her.
 *
 * Every link between a filed document and a horse is her name in the
 * filename. Rename her and that link breaks for everything already on
 * file: the vet bill that says "New recip" stops belonging to anybody the
 * moment she becomes Appaloosa Recip, and her care history and her spend
 * quietly drop off the back.
 *
 * Nobody should have to maintain a list of old names, so the app records
 * one for itself. A name is only ever written here at the moment it stops
 * being current — on a rename or a merge — which is rare, so this file is
 * read often and written almost never.
 */
export type Aliases = Record<string, string[]>;

const clean = (s: unknown) =>
  typeof s === "string" && s.trim().length >= 2 ? s.trim() : null;

export async function readAliases(): Promise<Aliases> {
  try {
    const p = encodeURIComponent(PATH()).replace(/%2F/g, "/");
    const res = await graphRaw(`/drives/${DRIVE}/root:/${p}:/content`);
    const data = JSON.parse(await res.text());
    if (data && typeof data === "object" && !Array.isArray(data)) {
      return data as Aliases;
    }
  } catch {
    /* none yet, or unreadable — a horse simply has no former names */
  }
  return {};
}

/**
 * Record names `id` used to go by. Returns quietly on failure: a lost
 * alias costs some history, a thrown error costs the save itself.
 */
export async function rememberNames(id: string, names: (string | null)[]) {
  const add = names.map(clean).filter((n): n is string => !!n);
  if (!add.length) return;

  try {
    const all = await readAliases();
    const have = all[id] ?? [];
    const seen = new Set(have.map((n) => n.toLowerCase()));

    let changed = false;
    for (const n of add) {
      if (seen.has(n.toLowerCase())) continue;
      seen.add(n.toLowerCase());
      have.push(n);
      changed = true;
    }
    if (!changed) return;

    all[id] = have;
    const p = encodeURIComponent(PATH()).replace(/%2F/g, "/");
    await graph(`/drives/${DRIVE}/root:/${p}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(all, null, 1),
    });
  } catch {
    /* best effort */
  }
}
