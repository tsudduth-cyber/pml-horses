import { graph, graphAll } from "@/lib/graph";
import { LIST_TEMPLATE } from "@/lib/schema";
import seed from "@/lib/seed.json";

export const dynamic = "force-dynamic";

const HOST = process.env.SP_HOSTNAME!;
const LIST_NAME = process.env.SP_LIST_NAME ?? "PML Horses";

/**
 * One-time setup. Creates the SharePoint list and loads the 26 horses.
 *
 * Run it once, signed in, from a browser:
 *     /api/setup?confirm=yes
 *
 * Safe to run twice. It will not create a second list and it will not
 * re-seed a list that already has rows.
 */
export async function GET(req: Request) {
  const url = new URL(req.url);
  const log: string[] = [];
  const say = (s: string) => log.push(s);

  if (url.searchParams.get("confirm") !== "yes") {
    return text(
      [
        "This creates the SharePoint list and loads the horses.",
        "",
        `  list   : ${LIST_NAME}`,
        `  site   : ${HOST}`,
        `  records: ${(seed as unknown[]).length}`,
        "",
        "Nothing has happened yet. To go ahead, add ?confirm=yes to this URL.",
      ].join("\n")
    );
  }

  try {
    const site: any = await graph(`/sites/${HOST}`);
    say(`Site: ${site.displayName ?? HOST}`);

    // 1 ── find or create the list
    const lists = await graphAll<any>(
      `/sites/${site.id}/lists?$select=id,name,displayName`
    );
    let list = lists.find(
      (l) =>
        l.displayName?.toLowerCase() === LIST_NAME.toLowerCase() ||
        l.name?.toLowerCase() === LIST_NAME.toLowerCase()
    );

    if (list) {
      say(`List "${LIST_NAME}" already exists — leaving it alone.`);
    } else {
      list = await graph(`/sites/${site.id}/lists`, {
        method: "POST",
        body: JSON.stringify(LIST_TEMPLATE),
      });
      say(
        `Created list "${LIST_NAME}" with ${LIST_TEMPLATE.columns.length} columns.`
      );
    }

    // 2 ── seed, but never on top of rows that are already there
    const existing = await graphAll<any>(
      `/sites/${site.id}/lists/${list.id}/items?$select=id&$top=5`
    );
    if (existing.length > 0) {
      say(
        `List already holds ${existing.length >= 5 ? "5+" : existing.length} rows — skipping the load.`
      );
      say("Delete the rows first if you want to start clean.");
      return text(log.join("\n"));
    }

    let ok = 0;
    const failed: string[] = [];
    // One at a time: 26 rows is nothing, and a failure names its own horse.
    for (const fields of seed as Record<string, unknown>[]) {
      const who =
        (fields.BarnName as string) ||
        (fields.Title as string) ||
        (fields.IdLine as string) ||
        "unnamed";
      try {
        await graph(`/sites/${site.id}/lists/${list.id}/items`, {
          method: "POST",
          body: JSON.stringify({ fields }),
        });
        ok++;
      } catch (e: any) {
        failed.push(`${who}: ${String(e?.message ?? e).slice(0, 160)}`);
      }
    }

    say(`Loaded ${ok} of ${(seed as unknown[]).length} horses.`);
    if (failed.length) {
      say("");
      say("These did not go in:");
      failed.forEach((f) => say("  " + f));
    } else {
      say("");
      say("Done. Open the app and the horses should be there.");
    }

    return text(log.join("\n"));
  } catch (e: any) {
    const msg = String(e?.message ?? e);
    if (msg === "NOT_SIGNED_IN") {
      return text("Sign in first, then come back to this URL.", 401);
    }
    say("");
    say("Stopped: " + msg.slice(0, 500));
    return text(log.join("\n"), 500);
  }
}

function text(body: string, status = 200) {
  return new Response(body, {
    status,
    headers: { "Content-Type": "text/plain; charset=utf-8" },
  });
}
