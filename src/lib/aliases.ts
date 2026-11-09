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
 * Replace the names `id` also answers to.
 *
 * This is the one the office drives, from the horse's own page. Vets and
 * farriers write what they write — "Sorrel Recip", "recip mare 1", a barn
 * name nobody told them was wrong — and every one of those is a bill the
 * app cannot place. Typing the name here is the whole fix: the scanner
 * reads these alongside her registered and barn names, in filenames and
 * inside the documents, so the next scan places the bill. It beats
 * waiting on a code change for what is really a spelling problem.
 *
 * Unlike rememberNames this can remove a name, because a wrong one typed
 * in here would quietly credit another horse's paperwork to this one.
 */
export async function setNames(id: string, names: string[]) {
  const next: string[] = [];
  const seen = new Set<string>();
  for (const raw of names) {
    const n = clean(raw);
    if (!n || seen.has(n.toLowerCase())) continue;
    seen.add(n.toLowerCase());
    next.push(n);
    if (next.length >= 20) break;
  }

  const all = await readAliases();
  if (next.length) all[id] = next;
  else delete all[id];

  const p = encodeURIComponent(PATH()).replace(/%2F/g, "/");
  await graph(`/drives/${DRIVE}/root:/${p}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(all, null, 1),
  });
  return next;
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
