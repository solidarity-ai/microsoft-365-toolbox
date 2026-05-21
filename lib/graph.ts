const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export interface GraphPage<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export interface GraphQueryValue {
  key: string;
  value: string | number | boolean | undefined;
}

export function graphUrl(path: string, query: GraphQueryValue[] = []): string {
  const normalizedPath = path.startsWith("/") ? path : `/${path}`;
  const params: string[] = [];
  for (const item of query) {
    if (item.value !== undefined) {
      params.push(`${encodeURIComponent(item.key)}=${encodeURIComponent(String(item.value))}`);
    }
  }
  return `${GRAPH_BASE_URL}${normalizedPath}${params.length > 0 ? `?${params.join("&")}` : ""}`;
}

export async function graphGet<T>(
  path: string,
  query: GraphQueryValue[] = [],
  headers: Record<string, string> = {},
): Promise<T> {
  const resp = await fetch(graphUrl(path, query), {
    headers: {
      Accept: "application/json",
      ...headers,
    },
  });

  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Microsoft Graph ${resp.status}: ${body}`);
  }

  return await resp.json() as T;
}

export async function graphGetBytes(
  path: string,
  query: GraphQueryValue[] = [],
  headers: Record<string, string> = {},
): Promise<Uint8Array> {
  const resp = await fetch(graphUrl(path, query), {
    headers: {
      Accept: "application/octet-stream",
      ...headers,
    },
  });

  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Microsoft Graph ${resp.status}: ${body}`);
  }

  return new Uint8Array(await resp.arrayBuffer());
}

export async function graphPost<T>(
  path: string,
  body: unknown = {},
  query: GraphQueryValue[] = [],
  headers: Record<string, string> = {},
): Promise<T> {
  const resp = await fetch(graphUrl(path, query), {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
      ...headers,
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Microsoft Graph ${resp.status}: ${text}`);
  }

  if (resp.status === 204) {
    return {} as T;
  }
  return await resp.json() as T;
}

export async function graphPatch<T>(
  path: string,
  body: unknown,
  query: GraphQueryValue[] = [],
  headers: Record<string, string> = {},
): Promise<T> {
  const resp = await fetch(graphUrl(path, query), {
    method: "PATCH",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
      ...headers,
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Microsoft Graph ${resp.status}: ${text}`);
  }

  if (resp.status === 204) {
    return {} as T;
  }
  return await resp.json() as T;
}

export async function graphGetAllPages<T>(
  path: string,
  query: GraphQueryValue[] = [],
  max_pages = 1,
): Promise<GraphPage<T>> {
  if (!Number.isInteger(max_pages) || max_pages < 1 || max_pages > 20) {
    throw new Error("max_pages must be an integer from 1 to 20");
  }

  let nextUrl: string | undefined = graphUrl(path, query);
  const value: T[] = [];
  let lastNextLink: string | undefined;

  for (let page = 0; page < max_pages && nextUrl; page++) {
    const resp = await fetch(nextUrl, {
      headers: {
        Accept: "application/json",
      },
    });

    if (!resp.ok) {
      const body = await resp.text();
      throw new Error(`Microsoft Graph ${resp.status}: ${body}`);
    }

    const data = await resp.json() as GraphPage<T>;
    if (Array.isArray(data.value)) {
      value.push(...data.value);
    }
    lastNextLink = data["@odata.nextLink"];
    nextUrl = lastNextLink;
  }

  return lastNextLink ? { value, "@odata.nextLink": lastNextLink } : { value };
}
