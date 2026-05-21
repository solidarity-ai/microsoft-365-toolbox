export interface ZipEntrySummary {
  name: string;
  compressed_size: number;
  uncompressed_size: number;
  crc32: string;
}

function u16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function u32(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
}

function hex32(value: number): string {
  return value.toString(16).padStart(8, "0");
}

export function listZipEntries(bytes: Uint8Array): ZipEntrySummary[] {
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 66000); i--) {
    if (u32(bytes, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) {
    throw new Error("file is not a readable ZIP/OpenXML package");
  }
  const total = u16(bytes, eocd + 10);
  let offset = u32(bytes, eocd + 16);
  const decoder = new TextDecoder();
  const entries: ZipEntrySummary[] = [];
  for (let i = 0; i < total; i++) {
    if (u32(bytes, offset) !== 0x02014b50) {
      throw new Error("invalid ZIP central directory");
    }
    const crc32 = u32(bytes, offset + 16);
    const compressed = u32(bytes, offset + 20);
    const uncompressed = u32(bytes, offset + 24);
    const nameLen = u16(bytes, offset + 28);
    const extraLen = u16(bytes, offset + 30);
    const commentLen = u16(bytes, offset + 32);
    const name = decoder.decode(bytes.slice(offset + 46, offset + 46 + nameLen));
    entries.push({
      name,
      compressed_size: compressed,
      uncompressed_size: uncompressed,
      crc32: hex32(crc32),
    });
    offset += 46 + nameLen + extraLen + commentLen;
  }
  return entries.sort((a, b) => a.name.localeCompare(b.name));
}

export function summarizeOpenXml(entries: ZipEntrySummary[]): Record<string, unknown> {
  const names = entries.map((entry) => entry.name);
  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  return {
    entry_count: entries.length,
    has_vba_project: byName.has("xl/vbaProject.bin"),
    vba_project_fingerprint: byName.get("xl/vbaProject.bin")?.crc32,
    chart_parts: names.filter((name) => name.startsWith("xl/charts/") && name.endsWith(".xml")),
    drawing_parts: names.filter((name) => name.startsWith("xl/drawings/")),
    pivot_parts: names.filter((name) => name.toLowerCase().includes("pivot")),
    table_parts: names.filter((name) => name.startsWith("xl/tables/") && name.endsWith(".xml")),
    external_link_parts: names.filter((name) => name.startsWith("xl/externalLinks/")),
    connection_parts: names.filter((name) => name === "xl/connections.xml" || name.startsWith("xl/connections/")),
    comments_parts: names.filter((name) => name.toLowerCase().includes("comment")),
    relationship_parts: names.filter((name) => name.endsWith(".rels")),
    styles_fingerprint: byName.get("xl/styles.xml")?.crc32,
    workbook_rels_fingerprint: byName.get("xl/_rels/workbook.xml.rels")?.crc32,
  };
}
