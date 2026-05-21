import { assertId, getRange, listWorksheets, summarizeFormulaPatterns, usedRange } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  worksheet_id_or_name?: string,
  range_address?: string,
  workbook_session_id?: string,
  include_risky = true,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const worksheets = worksheet_id_or_name
    ? [{ name: worksheet_id_or_name }]
    : await listWorksheets(drive_id, item_id, workbook_session_id);
  const sheets = [];
  let formula_count = 0;
  const risky = [];
  const unique = new Set<string>();
  for (const ws of worksheets) {
    const range = range_address
      ? await getRange(drive_id, item_id, ws.name, range_address, workbook_session_id, "address,rowCount,columnCount,formulas")
      : await usedRange(drive_id, item_id, ws.name, workbook_session_id, false);
    const summary = summarizeFormulaPatterns(range);
    formula_count += summary.formula_count;
    for (const pattern of summary.unique_patterns) unique.add(pattern);
    if (include_risky) {
      risky.push(...summary.risky.map((item) => ({ ...item, sheet: ws.name })));
    }
    sheets.push({ sheet: ws.name, address: range.address, formula_count: summary.formula_count });
  }
  return { formula_count, unique_patterns: [...unique].slice(0, 100), risky, sheets };
}
