import { graphRaw } from "@/lib/graph";

export const dynamic = "force-dynamic";

const DRIVE = process.env.SP_DRIVE_ID!;

/**
 * Serve a horse photo through the app.
 *
 * The obvious thing — pointing <img src> straight at the SharePoint webUrl —
 * only works for a viewer whose browser already holds a SharePoint session.
 * On a phone it does not, and every photo renders as a broken image. So the
 * image is fetched here with the app's own token and streamed back.
 *
 *   /api/photo?id=<driveItemId>           full size
 *   /api/photo?id=<driveItemId>&size=sm   thumbnail, for lists
 */
export async function GET(req: Request) {
  const url = new URL(req.url);
  const id = url.searchParams.get("id") ?? "";
  const small = url.searchParams.get("size") === "sm";

  // Drive item ids are opaque but bounded; reject anything that isn't one
  // rather than passing arbitrary text into a Graph path.
  if (!/^[A-Za-z0-9!._-]{1,200}$/.test(id)) {
    return new Response("bad id", { status: 400 });
  }

  const path = small
    ? `/drives/${DRIVE}/items/${id}/thumbnails/0/large/content`
    : `/drives/${DRIVE}/items/${id}/content`;

  try {
    const res = await graphRaw(path);
    const body = await res.arrayBuffer();
    return new Response(body, {
      status: 200,
      headers: {
        "Content-Type": res.headers.get("Content-Type") ?? "image/jpeg",
        // Private: these are Philip's horses, not public images. Caching in
        // the viewer's own browser is what makes scrolling the list quick.
        "Cache-Control": "private, max-age=3600",
      },
    });
  } catch (e: any) {
    const msg = String(e?.message ?? e);
    if (msg === "NOT_SIGNED_IN") return new Response("sign in", { status: 401 });
    // A missing thumbnail is normal for some file types — fall back to the
    // full image rather than showing a broken one.
    if (small) {
      try {
        const res = await graphRaw(`/drives/${DRIVE}/items/${id}/content`);
        const body = await res.arrayBuffer();
        return new Response(body, {
          status: 200,
          headers: {
            "Content-Type": res.headers.get("Content-Type") ?? "image/jpeg",
            "Cache-Control": "private, max-age=3600",
          },
        });
      } catch {
        /* fall through */
      }
    }
    return new Response("not available", { status: 502 });
  }
}
