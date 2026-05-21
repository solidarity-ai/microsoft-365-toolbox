import { assertId, listWorksheets, summarizeRangeErrors, usedRange } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  workbook_session_id?: string,
  worksheet_id_or_name?: string,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const worksheets = worksheet_id_or_name
    ? [{ name: worksheet_id_or_name }]
    : await listWorksheets(drive_id, item_id, workbook_session_id);
  const errors = [];
  for (const ws of worksheets) {
    const range = await usedRange(drive_id, item_id, ws.name, workbook_session_id, false);
    errors.push(...summarizeRangeErrors(range).map((item) => ({ ...item, sheet: ws.name, range: range.address })));
  }
  return { error_count: errors.length, errors: errors.slice(0, 200), truncated: errors.length > 200 };
}
