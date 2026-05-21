import { assertId, getRange } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  worksheet_id_or_name: string,
  address: string,
  workbook_session_id?: string,
  include_values = true,
  include_formulas = false,
  include_number_formats = false,
  include_text = false,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const fields = ["address", "rowCount", "columnCount"];
  if (include_values) fields.push("values");
  if (include_formulas) fields.push("formulas");
  if (include_number_formats) fields.push("numberFormat");
  if (include_text) fields.push("text");
  return await getRange(drive_id, item_id, worksheet_id_or_name, address, workbook_session_id, fields.join(","));
}
