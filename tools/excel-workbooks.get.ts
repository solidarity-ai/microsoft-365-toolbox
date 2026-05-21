import { graphGet } from "../lib/graph.ts";
import { assertId, type GraphDriveItem } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const item = await graphGet<GraphDriveItem>(
    `/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}`,
  );
  const extension = item.name?.split(".").pop()?.toLowerCase() ?? "";
  const graph_excel_supported = extension === "xlsx";
  return {
    ...item,
    graph_excel_supported,
    macro_enabled: extension === "xlsm" || extension === "xlsb",
    recommended_backend: graph_excel_supported ? "graph" : "download_parse",
  };
}
