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
  baseline_values: unknown[][],
  workbook_session_id?: string,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  if (!Array.isArray(baseline_values)) {
    throw new Error("baseline_values must be a 2D array captured before preview edits");
  }
  const current = await getRange(drive_id, item_id, worksheet_id_or_name, address, workbook_session_id, "address,values,formulas,numberFormat");
  const values = current.values ?? [];
  const changed = [];
  const maxRows = Math.max(values.length, baseline_values.length);
  for (let r = 0; r < maxRows; r++) {
    const oldRow = baseline_values[r] ?? [];
    const newRow = values[r] ?? [];
    const maxCols = Math.max(oldRow.length, newRow.length);
    for (let c = 0; c < maxCols; c++) {
      if (JSON.stringify(oldRow[c]) !== JSON.stringify(newRow[c])) {
        changed.push({ row: r + 1, column: c + 1, before: oldRow[c], after: newRow[c] });
      }
    }
  }
  return { address: current.address, changed_cell_count: changed.length, changed_cells: changed.slice(0, 200), truncated: changed.length > 200 };
}
