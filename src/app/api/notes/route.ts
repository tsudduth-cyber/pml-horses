import { NextResponse } from "next/server";
import { saveLines } from "@/lib/sharepoint";
import type { TextKey } from "@/lib/sharepoint";

const ALLOWED: TextKey[] = ["notes", "missing"];

export async function POST(req: Request) {
  let body: any;
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "Bad request" }, { status: 400 });
  }

  const { id, field, lines } = body ?? {};

  if (typeof id !== "string" || !id) {
    return NextResponse.json({ error: "Missing horse" }, { status: 400 });
  }
  if (!ALLOWED.includes(field)) {
    return NextResponse.json({ error: "Not an editable field" }, { status: 400 });
  }
  if (!Array.isArray(lines) || lines.some((l) => typeof l !== "string")) {
    return NextResponse.json(
      { error: "Notes must be lines of text" },
      { status: 400 }
    );
  }

  try {
    const saved = await saveLines(id, field, lines);
    return NextResponse.json({ ok: true, lines: saved });
  } catch (e: any) {
    if (e?.message === "NOT_SIGNED_IN") {
      return NextResponse.json({ error: "Sign in again" }, { status: 401 });
    }
    if (e?.message === "TOO_LONG") {
      return NextResponse.json({ error: "That note is too long" }, { status: 400 });
    }
    return NextResponse.json(
      { error: "SharePoint wouldn't take it" },
      { status: 502 }
    );
  }
}
