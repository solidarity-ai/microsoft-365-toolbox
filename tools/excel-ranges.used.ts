import { assertId, usedRange } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  worksheet_id_or_name: string,
  workbook_session_id?: string,
  values_only = false,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  assertId("worksheet_id_or_name", worksheet_id_or_name);
  return await usedRange(drive_id, item_id, worksheet_id_or_name, workbook_session_id, values_only);
}
