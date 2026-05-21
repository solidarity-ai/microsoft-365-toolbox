import { graphGet } from "../lib/graph.ts";

interface GraphDriveItem {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: unknown;
  folder?: unknown;
  package?: unknown;
  parentReference?: unknown;
}

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
): Promise<unknown> {
  if (typeof drive_id !== "string" || drive_id.trim() === "") {
    throw new Error("drive_id is required");
  }
  if (typeof item_id !== "string" || item_id.trim() === "") {
    throw new Error("item_id is required");
  }

  return await graphGet<GraphDriveItem>(
    `/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}`,
  );
}
