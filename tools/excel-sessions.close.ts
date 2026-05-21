import { graphPost } from "../lib/graph.ts";
import { assertId, workbookPath, workbookSessionHeaders } from "../lib/excel.ts";

/**
 * @effect reversible
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  workbook_session_id: string,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  assertId("workbook_session_id", workbook_session_id);
  await graphPost(
    workbookPath(drive_id, item_id, "/closeSession"),
    {},
    [],
    workbookSessionHeaders(workbook_session_id),
  );
  return { status: "closed" };
}
