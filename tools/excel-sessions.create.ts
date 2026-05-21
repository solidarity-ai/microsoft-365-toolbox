import { assertId, workbookCreateSession } from "../lib/excel.ts";

/**
 * @effect reversible
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  persist_changes = false,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  if (typeof persist_changes !== "boolean") {
    throw new Error("persist_changes must be boolean");
  }
  const session = await workbookCreateSession(drive_id, item_id, persist_changes);
  return {
    workbook_session_id: session.id,
    persist_changes,
    status: session.id ? "created" : "created_without_id",
  };
}
