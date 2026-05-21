import { graphGetBytes } from "../lib/graph.ts";
import { assertId } from "../lib/excel.ts";
import { listZipEntries, summarizeOpenXml, type ZipEntrySummary } from "../lib/zip.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  baseline_entries: ZipEntrySummary[],
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  if (!Array.isArray(baseline_entries)) {
    throw new Error("baseline_entries must be entries returned by excelWorkbooks.packageInspect(include_entries=true)");
  }
  const bytes = await graphGetBytes(`/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}/content`);
  const current = listZipEntries(bytes);
  const oldMap = new Map(baseline_entries.map((entry) => [entry.name, entry]));
  const newMap = new Map(current.map((entry) => [entry.name, entry]));
  const added = current.filter((entry) => !oldMap.has(entry.name)).map((entry) => entry.name);
  const removed = baseline_entries.filter((entry) => !newMap.has(entry.name)).map((entry) => entry.name);
  const changed = current
    .filter((entry) => oldMap.has(entry.name) && oldMap.get(entry.name)?.crc32 !== entry.crc32)
    .map((entry) => entry.name);
  const oldSummary = summarizeOpenXml(baseline_entries);
  const newSummary = summarizeOpenXml(current);
  return {
    added_parts: added,
    removed_parts: removed,
    changed_parts: changed,
    before: oldSummary,
    after: newSummary,
    vba_project_preserved: oldSummary.vba_project_fingerprint === newSummary.vba_project_fingerprint,
    styles_preserved: oldSummary.styles_fingerprint === newSummary.styles_fingerprint,
    workbook_relationships_preserved: oldSummary.workbook_rels_fingerprint === newSummary.workbook_rels_fingerprint,
  };
}
