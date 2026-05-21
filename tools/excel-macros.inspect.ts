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
    has_vba_project: summary.has_vba_project,
    vba_project_fingerprint: summary.vba_project_fingerprint,
    macro_mutation_supported: false,
    safe_handling: summary.has_vba_project
      ? "Inspect and preserve byte-for-byte; refuse persistent mutation unless a macro-preserving path is available."
      : "No VBA project detected.",
  };
}
