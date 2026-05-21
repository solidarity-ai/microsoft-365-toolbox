import { graphGetBytes } from "../lib/graph.ts";
import { assertId, listWorksheets, summarizeFormulaPatterns, usedRange } from "../lib/excel.ts";
import { listZipEntries, summarizeOpenXml } from "../lib/zip.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  workbook_session_id?: string,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const bytes = await graphGetBytes(`/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}/content`);
  const packageSummary = summarizeOpenXml(listZipEntries(bytes));
  const risky_formulas = [];
  for (const ws of await listWorksheets(drive_id, item_id, workbook_session_id)) {
    const range = await usedRange(drive_id, item_id, ws.name, workbook_session_id, false);
    risky_formulas.push(...summarizeFormulaPatterns(range).risky.map((item) => ({ ...item, sheet: ws.name })));
  }
  return {
    external_link_parts: packageSummary.external_link_parts,
    connection_parts: packageSummary.connection_parts,
    risky_formulas: risky_formulas.filter((item) => item.reason === "external_reference" || item.reason === "WEBSERVICE" || item.reason === "RTD"),
    refresh_executed: false,
  };
}
