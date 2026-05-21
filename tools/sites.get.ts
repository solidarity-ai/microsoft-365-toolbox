import { graphGet } from "../lib/graph.ts";

interface GraphSite {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  site_id: string,
): Promise<unknown> {
  if (typeof site_id !== "string" || site_id.trim() === "") {
    throw new Error("site_id is required");
  }

  return await graphGet<GraphSite>(`/sites/${encodeURIComponent(site_id)}`);
}
