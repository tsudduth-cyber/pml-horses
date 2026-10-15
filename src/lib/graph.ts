import { getAccessToken } from "./auth";

const GRAPH = "https://graph.microsoft.com/v1.0";

export class GraphError extends Error {
  constructor(public status: number, public detail: string) {
    super(`Graph ${status}: ${detail}`);
  }
}

export async function graph<T = any>(
  path: string,
  init: RequestInit = {}
): Promise<T> {
  const token = await getAccessToken();
  const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init.headers ?? {}),
    },
    cache: "no-store",
  });

  if (!res.ok) {
    throw new GraphError(res.status, (await res.text()).slice(0, 400));
  }
  if (res.status === 204) return undefined as T;
  return res.json();
}

/** Follows @odata.nextLink so a long list comes back whole. */
export async function graphAll<T = any>(path: string): Promise<T[]> {
  const out: T[] = [];
  let next: string | undefined = path;
  while (next) {
    const page: any = await graph(next);
    out.push(...(page.value ?? []));
    next = page["@odata.nextLink"];
  }
  return out;
}
