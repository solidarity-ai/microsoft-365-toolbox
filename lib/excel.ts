import { graphGet, graphPatch, graphPost, type GraphQueryValue } from "./graph.ts";

export interface GraphDriveItem {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string };
  folder?: unknown;
  parentReference?: {
    driveId?: string;
    path?: string;
  };
}

export interface GraphCollection<T> {
  value: T[];
}

export interface Worksheet {
  id: string;
  name: string;
  position?: number;
  visibility?: string;
}

export interface WorkbookRange {
  address?: string;
  addressLocal?: string;
  rowCount?: number;
  columnCount?: number;
  cellCount?: number;
  values?: unknown[][];
  formulas?: unknown[][];
  formulasLocal?: unknown[][];
  text?: string[][];
  numberFormat?: unknown[][];
  valueTypes?: string[][];
}

export function workbookSessionHeaders(workbook_session_id?: string): Record<string, string> {
  return workbook_session_id ? { "workbook-session-id": workbook_session_id } : {};
}

export function workbookPath(drive_id: string, item_id: string, suffix = ""): string {
  assertId("drive_id", drive_id);
  assertId("item_id", item_id);
  return `/drives/${encodeURIComponent(drive_id)}/items/${encodeURIComponent(item_id)}/workbook${suffix}`;
}

export function assertId(name: string, value: string): void {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${name} is required`);
  }
}

export function assertAddress(address: string): void {
  if (typeof address !== "string" || address.trim() === "") {
    throw new Error("address is required");
  }
  if (address.length > 120 || /[;\n\r]/.test(address)) {
    throw new Error("address must be a single A1 range");
  }
}

export function worksheetItemPath(worksheet_id_or_name: string): string {
  assertId("worksheet_id_or_name", worksheet_id_or_name);
  const escaped = worksheet_id_or_name.replaceAll("'", "''");
  return `worksheets('${escaped}')`;
}

export function rangeFunction(address: string): string {
  assertAddress(address);
  return `range(address='${address.replaceAll("'", "''")}')`;
}

export async function getWorksheet(
  drive_id: string,
  item_id: string,
  worksheet_id_or_name: string,
  workbook_session_id?: string,
): Promise<Worksheet> {
  return await graphGet<Worksheet>(
    workbookPath(drive_id, item_id, `/${worksheetItemPath(worksheet_id_or_name)}`),
    [],
    workbookSessionHeaders(workbook_session_id),
  );
}

export async function listWorksheets(
  drive_id: string,
  item_id: string,
  workbook_session_id?: string,
): Promise<Worksheet[]> {
  const data = await graphGet<GraphCollection<Worksheet>>(
    workbookPath(drive_id, item_id, "/worksheets"),
    [],
    workbookSessionHeaders(workbook_session_id),
  );
  return data.value;
}

export async function usedRange(
  drive_id: string,
  item_id: string,
  worksheet: string,
  workbook_session_id?: string,
  values_only = false,
): Promise<WorkbookRange> {
  const query: GraphQueryValue[] = values_only
    ? [{ key: "$select", value: "address,rowCount,columnCount,values,text" }]
    : [];
  return await graphGet<WorkbookRange>(
    workbookPath(drive_id, item_id, `/${worksheetItemPath(worksheet)}/usedRange()`),
    query,
    workbookSessionHeaders(workbook_session_id),
  );
}

export async function getRange(
  drive_id: string,
  item_id: string,
  worksheet: string,
  address: string,
  workbook_session_id?: string,
  select?: string,
): Promise<WorkbookRange> {
  return await graphGet<WorkbookRange>(
    workbookPath(drive_id, item_id, `/${worksheetItemPath(worksheet)}/${rangeFunction(address)}`),
    select ? [{ key: "$select", value: select }] : [],
    workbookSessionHeaders(workbook_session_id),
  );
}

export async function patchRange(
  drive_id: string,
  item_id: string,
  worksheet: string,
  address: string,
  body: Record<string, unknown>,
  workbook_session_id?: string,
): Promise<WorkbookRange> {
  return await graphPatch<WorkbookRange>(
    workbookPath(drive_id, item_id, `/${worksheetItemPath(worksheet)}/${rangeFunction(address)}`),
    body,
    [],
    workbookSessionHeaders(workbook_session_id),
  );
}

export async function workbookTables(
  drive_id: string,
  item_id: string,
  workbook_session_id?: string,
): Promise<unknown[]> {
  const data = await graphGet<GraphCollection<unknown>>(
    workbookPath(drive_id, item_id, "/tables"),
    [],
    workbookSessionHeaders(workbook_session_id),
  );
  return data.value;
}

export async function workbookNames(
  drive_id: string,
  item_id: string,
  workbook_session_id?: string,
): Promise<unknown[]> {
  const data = await graphGet<GraphCollection<unknown>>(
    workbookPath(drive_id, item_id, "/names"),
    [],
    workbookSessionHeaders(workbook_session_id),
  );
  return data.value;
}

export async function workbookCreateSession(
  drive_id: string,
  item_id: string,
  persist_changes: boolean,
): Promise<{ id?: string; persistChanges?: boolean }> {
  return await graphPost<{ id?: string; persistChanges?: boolean }>(
    workbookPath(drive_id, item_id, "/createSession"),
    { persistChanges: persist_changes },
  );
}

export async function workbookCalculate(
  drive_id: string,
  item_id: string,
  calculation_type: string,
  workbook_session_id?: string,
): Promise<unknown> {
  return await graphPost<unknown>(
    workbookPath(drive_id, item_id, "/application/calculate"),
    { calculationType: calculation_type },
    [],
    workbookSessionHeaders(workbook_session_id),
  );
}

export function summarizeRangeErrors(range: WorkbookRange): Array<{ row: number; column: number; value: unknown }> {
  const values = range.values ?? range.text ?? [];
  const errors = new Set(["#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"]);
  const out: Array<{ row: number; column: number; value: unknown }> = [];
  for (let r = 0; r < values.length; r++) {
    const row = values[r] ?? [];
    for (let c = 0; c < row.length; c++) {
      if (errors.has(String(row[c]))) {
        out.push({ row: r + 1, column: c + 1, value: row[c] });
      }
    }
  }
  return out;
}

export function summarizeFormulaPatterns(range: WorkbookRange): {
  formula_count: number;
  unique_patterns: string[];
  risky: Array<{ row: number; column: number; formula: string; reason: string }>;
} {
  const formulas = range.formulas ?? [];
  const patterns = new Set<string>();
  const risky: Array<{ row: number; column: number; formula: string; reason: string }> = [];
  const riskyFns = ["WEBSERVICE", "HYPERLINK", "INDIRECT", "RTD", "IMPORTXML"];
  let count = 0;
  for (let r = 0; r < formulas.length; r++) {
    const row = formulas[r] ?? [];
    for (let c = 0; c < row.length; c++) {
      const formula = String(row[c] ?? "");
      if (!formula.startsWith("=")) continue;
      count++;
      patterns.add(formula.replace(/\$?[A-Z]{1,3}\$?\d+/g, "REF").replace(/\d+(?:\.\d+)?/g, "N"));
      const upper = formula.toUpperCase();
      for (const fn of riskyFns) {
        if (upper.includes(`${fn}(`)) {
          risky.push({ row: r + 1, column: c + 1, formula, reason: fn });
        }
      }
      if (/\[[^\]]+\]/.test(formula)) {
        risky.push({ row: r + 1, column: c + 1, formula, reason: "external_reference" });
      }
    }
  }
  return { formula_count: count, unique_patterns: [...patterns].slice(0, 100), risky };
}
