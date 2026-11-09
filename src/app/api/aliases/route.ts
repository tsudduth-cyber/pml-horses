import { NextResponse } from "next/server";
import { setNames } from "@/lib/aliases";
import { forgetScan } from "@/lib/sharepoint";

/**
 * The other names a horse answers to.
 *
 * Saved by the office from the horse's own page. A name added here is
 * read by the scanner on its next pass, so a bill that said "Sorrel
 * Recip" lands on the right horse without anyone touching the code.
 */
export async function POST(req: Request) {
  let body: any;
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "Bad request" }, { status: 400 });
  }

  const { id, names } = body ?? {};
  if (typeof id !== "string" || !id) {
    return NextResponse.json({ error: "Missing horse" }, { status: 400 });
  }
  if (!Array.isArray(names) || names.some((n) => typeof n !== "string")) {
    return NextResponse.json({ error: "Names must be text" }, { status: 400 });
  }
  if (names.some((n) => n.length > 120)) {
    return NextResponse.json({ error: "That name is too long" }, { status: 400 });
  }

  try {
    const saved = await setNames(id, names);
    // A new name changes who the existing paperwork belongs to, so the
    // saved reading is wrong rather than merely old.
    forgetScan();
    return NextResponse.json({ ok: true, names: saved });
  } catch (e: any) {
    if (e?.message === "NOT_SIGNED_IN") {
      return NextResponse.json({ error: "Sign in again" }, { status: 401 });
    }
    return NextResponse.json(
      { error: "SharePoint wouldn't take it" },
      { status: 502 }
    );
  }
}
