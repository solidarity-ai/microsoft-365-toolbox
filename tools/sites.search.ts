import { graphGetAllPages } from "../lib/graph.ts";

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
  query: string,
  max_pages = 1,
): Promise<unknown> {
  if (typeof query !== "string" || query.trim() === "") {
    throw new Error("query is required");
  }

  return await graphGetAllPages<GraphSite>(
    "/sites",
    [{ key: "search", value: query }],
    max_pages,
  );
}
