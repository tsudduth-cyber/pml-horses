import { NextResponse } from "next/server";
import { saveRecord } from "@/lib/sharepoint";
import type { RecordKey } from "@/lib/sharepoint";

const ALLOWED: RecordKey[] = [
  "barn",
  "reg",
  "aqha",
  "sex",
  "color",
  "dob",
  "bought",
  "folderName",
];

const MESSAGE: Record<string, string> = {
  BAD_DATE: "Dates need to be a real date",
  BAD_SEX: "Pick one of the listed sexes",
  TOO_LONG: "That is too long for the field",
  BAD_FIELD: "Not an editable field",
};

export async function POST(req: Request) {
  let body: any;
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "Bad request" }, { status: 400 });
  }

  const { id, fields } = body ?? {};

  if (typeof id !== "string" || !id) {
    return NextResponse.json({ error: "Missing horse" }, { status: 400 });
  }
  if (!fields || typeof fields !== "object") {
    return NextResponse.json({ error: "Nothing to save" }, { status: 400 });
  }
  const keys = Object.keys(fields) as RecordKey[];
  if (!keys.length || keys.some((k) => !ALLOWED.includes(k))) {
    return NextResponse.json({ error: "Not an editable field" }, { status: 400 });
  }
  if (keys.some((k) => fields[k] !== null && typeof fields[k] !== "string")) {
    return NextResponse.json({ error: "Bad value" }, { status: 400 });
  }

  try {
    await saveRecord(id, fields);
    return NextResponse.json({ ok: true });
  } catch (e: any) {
    const m = String(e?.message ?? e);
    if (m === "NOT_SIGNED_IN") {
      return NextResponse.json({ error: "Sign in again" }, { status: 401 });
    }
    if (MESSAGE[m]) return NextResponse.json({ error: MESSAGE[m] }, { status: 400 });
    return NextResponse.json(
      { error: "SharePoint wouldn't take it" },
      { status: 502 }
    );
  }
}
