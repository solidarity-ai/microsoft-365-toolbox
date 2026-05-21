import { assertId, listWorksheets } from "../lib/excel.ts";

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
  return await listWorksheets(drive_id, item_id, workbook_session_id);
}
