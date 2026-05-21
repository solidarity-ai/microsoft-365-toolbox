import { assertId, workbookCalculate } from "../lib/excel.ts";

/**
 * @effect reversible
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  workbook_session_id?: string,
  calculation_type = "Full",
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const allowed = new Set(["Recalculate", "Full", "FullRebuild"]);
  if (!allowed.has(calculation_type)) {
    throw new Error("calculation_type must be Recalculate, Full, or FullRebuild");
  }
  await workbookCalculate(drive_id, item_id, calculation_type, workbook_session_id);
  return { status: "calculated", calculation_type };
}
