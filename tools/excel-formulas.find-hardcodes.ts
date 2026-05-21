import { assertId, getRange, listWorksheets, usedRange } from "../lib/excel.ts";

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  drive_id: string,
  item_id: string,
  worksheet_id_or_name?: string,
  range_address?: string,
  workbook_session_id?: string,
): Promise<unknown> {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  const worksheets = worksheet_id_or_name
    ? [{ name: worksheet_id_or_name }]
    : await listWorksheets(drive_id, item_id, workbook_session_id);
  const candidates = [];
  for (const ws of worksheets) {
    const range = range_address
      ? await getRange(drive_id, item_id, ws.name, range_address, workbook_session_id, "address,values,formulas")
      : await usedRange(drive_id, item_id, ws.name, workbook_session_id, false);
    const values = range.values ?? [];
    const formulas = range.formulas ?? [];
    for (let r = 0; r < values.length; r++) {
      const row = values[r] ?? [];
      const formulaRow = formulas[r] ?? [];
      for (let c = 0; c < row.length; c++) {
        const formula = String(formulaRow[c] ?? "");
        if (!formula.startsWith("=") && typeof row[c] === "number") {
          candidates.push({ sheet: ws.name, row: r + 1, column: c + 1, value: row[c], reason: "numeric_constant_without_formula" });
        }
      }
    }
  }
  return { candidate_count: candidates.length, candidates: candidates.slice(0, 200), truncated: candidates.length > 200 };
}
