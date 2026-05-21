import { graphGetBytes } from "../lib/graph.ts";
import { assertId } from "../lib/excel.ts";
import { listZipEntries, summarizeOpenXml } from "../lib/zip.ts";

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
  const bytes = await graphGetBytes(`/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}/content`);
  const summary = summarizeOpenXml(listZipEntries(bytes));
  return {
    chart_parts: summary.chart_parts,
    drawing_parts: summary.drawing_parts,
    relationship_parts: summary.relationship_parts,
    note: "OpenXML part inventory; use package diff after edits to verify chart and drawing preservation.",
  };
}
