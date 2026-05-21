import { assertId, listWorksheets, usedRange, summarizeFormulaPatterns } from "../lib/excel.ts";
import { graphGetBytes } from "../lib/graph.ts";
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
  const flags = new Set<string>();
  const bytes = await graphGetBytes(`/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}/content`);
  const packageSummary = summarizeOpenXml(listZipEntries(bytes));
  if (packageSummary.has_vba_project) flags.add("macro_enabled_workbook");
  if ((packageSummary.external_link_parts as unknown[]).length > 0) flags.add("external_link_parts");
  if ((packageSummary.connection_parts as unknown[]).length > 0) flags.add("workbook_connections");
  if ((packageSummary.pivot_parts as unknown[]).length > 0) flags.add("pivot_parts");
  if ((packageSummary.chart_parts as unknown[]).length > 0) flags.add("chart_parts");
  const risky_formulas = [];
  for (const ws of await listWorksheets(drive_id, item_id, workbook_session_id)) {
    if (ws.visibility && ws.visibility !== "Visible") flags.add("hidden_sheet");
    const range = await usedRange(drive_id, item_id, ws.name, workbook_session_id, false);
    const summary = summarizeFormulaPatterns(range);
    if (summary.risky.length > 0) {
      flags.add("risky_formula");
      risky_formulas.push(...summary.risky.map((item) => ({ ...item, sheet: ws.name })));
    }
  }
  return { risk_flags: [...flags].sort(), package: packageSummary, risky_formulas };
}
