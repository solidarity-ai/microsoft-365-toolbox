import { assertId, listWorksheets, usedRange, workbookNames, workbookTables } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  workbook_session_id?: string,
  include_hidden = false,
  include_samples = false,
  max_sample_rows = 5,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  if (!Number.isInteger(max_sample_rows) || max_sample_rows < 0 || max_sample_rows > 25) {
    throw new Error("max_sample_rows must be an integer from 0 to 25");
  }
  const worksheets = await listWorksheets(drive_id, item_id, workbook_session_id);
  const sheets = [];
  for (const ws of worksheets) {
    if (!include_hidden && ws.visibility && ws.visibility !== "Visible") {
      sheets.push({ id: ws.id, name: ws.name, visibility: ws.visibility, hidden: true });
      continue;
    }
    const range = await usedRange(drive_id, item_id, ws.name, workbook_session_id, !include_samples);
    const sample = include_samples && range.values
      ? range.values.slice(0, max_sample_rows)
      : undefined;
    sheets.push({
      id: ws.id,
      name: ws.name,
      position: ws.position,
      visibility: ws.visibility,
      used_range: {
        address: range.address,
        row_count: range.rowCount,
        column_count: range.columnCount,
      },
      sample,
    });
  }
  let tables: unknown[] = [];
  let names: unknown[] = [];
  try {
    tables = await workbookTables(drive_id, item_id, workbook_session_id);
  } catch (error) {
    tables = [{ error: error instanceof Error ? error.message : String(error) }];
  }
  try {
    names = await workbookNames(drive_id, item_id, workbook_session_id);
  } catch (error) {
    names = [{ error: error instanceof Error ? error.message : String(error) }];
  }
  return { worksheets: sheets, tables, named_items: names };
}
