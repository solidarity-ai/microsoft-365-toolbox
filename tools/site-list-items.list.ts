import { graphGetAllPages } from "../lib/graph.ts";

interface GraphListItem {
  id: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  fields?: unknown;
}

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  site_id: string,
  list_id: string,
  max_pages = 1,
): Promise<unknown> {
  if (typeof site_id !== "string" || site_id.trim() === "") {
    throw new Error("site_id is required");
  }
  if (typeof list_id !== "string" || list_id.trim() === "") {
    throw new Error("list_id is required");
  }

  return await graphGetAllPages<GraphListItem>(
    `/sites/${encodeURIComponent(site_id)}/lists/${encodeURIComponent(list_id)}/items`,
    [{ key: "$expand", value: "fields" }],
    max_pages,
  );
}
