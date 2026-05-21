import { graphGetAllPages } from "../lib/graph.ts";

interface GraphDrive {
  id: string;
  name?: string;
  driveType?: string;
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
  max_pages = 1,
): Promise<unknown> {
  if (typeof site_id !== "string" || site_id.trim() === "") {
    throw new Error("site_id is required");
  }

  return await graphGetAllPages<GraphDrive>(
    `/sites/${encodeURIComponent(site_id)}/drives`,
    [],
    max_pages,
  );
}
