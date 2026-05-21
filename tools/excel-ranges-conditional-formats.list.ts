import { assertId } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  worksheet_id_or_name: string,
  address?: string,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  assertId("worksheet_id_or_name", worksheet_id_or_name);
  return {
    worksheet: worksheet_id_or_name,
    address,
    supported_via_graph: false,
    recommendation: "Use packageInspect/packageDiff or local OpenXML checks to verify conditional formatting preservation.",
  };
}
