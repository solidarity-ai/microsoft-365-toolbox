import { graphGet } from "../lib/graph.ts";

interface SearchResponse {
  value: Array<{
    id: string;
    name?: string;
    webUrl?: string;
    size?: number;
    lastModifiedDateTime?: string;
    file?: { mimeType?: string };
    parentReference?: { driveId?: string; path?: string };
  }>;
}

function isExcel(name?: string): boolean {
  return /\.(xlsx|xlsm|xlsb)$/i.test(name ?? "");
}

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  query: string,
  drive_id?: string,
  max_items = 20,
): Promise<unknown> {
  if (typeof query !== "string" || query.trim() === "") {
    throw new Error("query is required");
  }
  if (!Number.isInteger(max_items) || max_items < 1 || max_items > 200) {
    throw new Error("max_items must be an integer from 1 to 200");
  }

  const escaped = query.replaceAll("'", "''");
  const path = drive_id
    ? `/drives/${encodeURIComponent(drive_id)}/root/search(q='${escaped}')`
    : `/me/drive/root/search(q='${escaped}')`;
  const data = await graphGet<SearchResponse>(path, [{ key: "$top", value: max_items }]);
  return data.value
    .filter((item) => isExcel(item.name))
    .slice(0, max_items)
    .map((item) => {
      const ext = item.name?.split(".").pop()?.toLowerCase();
      return {
        drive_id: item.parentReference?.driveId ?? drive_id,
        item_id: item.id,
        name: item.name,
        web_url: item.webUrl,
        size: item.size,
        last_modified: item.lastModifiedDateTime,
        path: item.parentReference?.path,
        file_extension: ext,
        graph_excel_supported: ext === "xlsx",
      };
    });
}
