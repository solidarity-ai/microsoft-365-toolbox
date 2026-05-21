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
  include_entries = false,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const bytes = await graphGetBytes(`/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}/content`);
  const entries = listZipEntries(bytes);
  return {
    byte_length: bytes.byteLength,
    ...summarizeOpenXml(entries),
    entries: include_entries ? entries : undefined,
  };
}
