import { NextResponse } from "next/server";
import { saveCareDate } from "@/lib/sharepoint";
import type { CareKey } from "@/lib/types";

const ALLOWED: CareKey[] = ["farrier", "worm", "vax", "coggins", "dental"];

export async function POST(req: Request) {
  let body: any;
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "Bad request" }, { status: 400 });
  }

  const { id, field, date } = body ?? {};

  if (typeof id !== "string" || !id) {
    return NextResponse.json({ error: "Missing horse" }, { status: 400 });
  }
  if (!ALLOWED.includes(field)) {
    return NextResponse.json({ error: "Not an editable field" }, { status: 400 });
  }
  if (date !== null && !/^\d{4}-\d{2}-\d{2}$/.test(String(date))) {
    return NextResponse.json({ error: "Date must be YYYY-MM-DD" }, { status: 400 });
  }

  try {
    await saveCareDate(id, field, date);
    return NextResponse.json({ ok: true });
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
