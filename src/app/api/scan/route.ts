import { getHorses, scanState } from "@/lib/sharepoint";
import { runScan } from "@/lib/scan";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

/**
 * Read the paperwork and record what it says.
 *
 * Called by the app when the last scan has gone stale, which means it runs
 * as whoever opened the app. Graph access here is delegated — a scheduled
 * job would have no one to act as — so the first person to open the app
 * after it goes stale pays for the refresh, and everyone after reads the
 * saved result.
 */
export async function POST(req: Request) {
  try {
    // A saved scan is only reread for four hours, which is right for normal
    // use. But when the reading itself changes — a new folder comes into
    // view, a keyword gets added — the saved answer is wrong rather than
    // merely old, and waiting out the clock means showing Philip dates we
    // already know are stale.
    const force = new URL(req.url).searchParams.get("force") === "1";

    const horses = await getHorses();
    const { stale, docs, contracts } = scanState();
    if (!stale && !force) return Response.json({ ok: true, skipped: "fresh" });

    const scan = await runScan(horses, docs, contracts);
    const found = Object.values(scan.items).reduce(
      (n, h) => n + Object.keys(h).length,
      0
    );
    return Response.json({ ok: true, at: scan.at, found });
  } catch (e: any) {
    const msg = String(e?.message ?? e);
    if (msg === "NOT_SIGNED_IN") {
      return Response.json({ error: "NOT_SIGNED_IN" }, { status: 401 });
    }
    return Response.json({ error: msg.slice(0, 200) }, { status: 502 });
  }
}
