import { assertId, patchRange } from "../lib/excel.ts";

/**
 * @effect reversible
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  worksheet_id_or_name: string,
  address: string,
  workbook_session_id: string,
  values?: unknown[][],
  formulas?: unknown[][],
  number_format?: unknown[][],
  preview_only = true,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  assertId("workbook_session_id", workbook_session_id);
  const body: Record<string, unknown> = {};
  if (values !== undefined) body.values = values;
  if (formulas !== undefined) body.formulas = formulas;
  if (number_format !== undefined) body.numberFormat = number_format;
  if (Object.keys(body).length === 0) {
    throw new Error("at least one of values, formulas, or number_format is required");
  }
  const range = await patchRange(drive_id, item_id, worksheet_id_or_name, address, body, workbook_session_id);
  return {
    status: preview_only ? "updated_in_session_preview" : "updated_in_session",
    preview_only,
    range,
  };
}
