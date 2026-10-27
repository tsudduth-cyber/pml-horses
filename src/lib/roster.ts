import type { Horse } from "./types";

/**
 * What a horse's folder name says about her.
 *
 * The filing convention is "Registered Name (Barn Name)", with a bare name
 * where there is only one, and a descriptive name while a horse is unnamed.
 * That convention is the roster: a folder is a horse, and renaming the
 * folder renames her.
 */
export function parseFolder(folder: string): {
  reg: string | null;
  barn: string | null;
  pending: boolean;
  idline: string | null;
} {
  const leaf = folder.split("/").pop()!.trim();

  // "Blazin Hot Metal (Sully)"
  const paren = leaf.match(/^(.+?)\s*\(([^)]+)\)\s*$/);
  if (paren) {
    return { reg: paren[1].trim(), barn: paren[2].trim(), pending: false, idline: null };
  }

  // "Lizard - Not Registering" — a note after the name, not a second name.
  const dash = leaf.match(/^(.+?)\s+-\s+(.+)$/);
  if (dash) {
    const [, name, note] = dash;
    // "Sue's Baby - Name TBD", "Unnamed Filly - Dual Reyish x High Brow Madonna"
    if (/name tbd|unnamed|tbd/i.test(leaf)) {
      return {
        reg: null,
        barn: null,
        pending: true,
        idline: /^unnamed/i.test(name) ? note.trim() : leaf,
      };
    }
    return { reg: null, barn: name.trim(), pending: false, idline: null };
  }

  if (/^unnamed/i.test(leaf)) {
    return { reg: null, barn: null, pending: true, idline: leaf };
  }

  return { reg: leaf, barn: null, pending: false, idline: null };
}

const key = (s: string | null | undefined) =>
  (s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");

/**
 * Attach a list row to a folder.
 *
 * The folder name first, because that is what the row records. Then the
 * names themselves, so that renaming a folder — which is how the office
 * renames a horse — keeps her price, status and care dates instead of
 * silently orphaning them.
 */
export function matchRow<T extends { folderName: string | null; barn: string | null; reg: string | null }>(
  folder: string,
  parsed: ReturnType<typeof parseFolder>,
  rows: T[],
  taken: Set<T>
): T | null {
  const tries: ((r: T) => boolean)[] = [
    (r) => key(r.folderName) === key(folder),
    (r) => !!parsed.barn && key(r.barn) === key(parsed.barn),
    (r) => !!parsed.reg && key(r.reg) === key(parsed.reg),
    (r) => !!parsed.reg && key(r.barn) === key(parsed.reg),
    (r) => !!parsed.idline && key(r.folderName) === key(parsed.idline),
  ];
  for (const t of tries) {
    const hit = rows.find((r) => !taken.has(r) && t(r));
    if (hit) return hit;
  }
  return null;
}
