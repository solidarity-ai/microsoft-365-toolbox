import { graphGet } from "../lib/graph.ts";

interface GraphDriveItem {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string; hashes?: unknown };
  parentReference?: {
    driveId?: string;
    driveType?: string;
    id?: string;
    name?: string;
    path?: string;
    siteId?: string;
  };
}

function encodeUtf8(value: string): number[] {
  const bytes: number[] = [];
  for (let i = 0; i < value.length; i++) {
    let codePoint = value.charCodeAt(i);
    if (codePoint >= 0xd800 && codePoint <= 0xdbff && i + 1 < value.length) {
      const next = value.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        codePoint = 0x10000 + ((codePoint - 0xd800) << 10) + (next - 0xdc00);
        i++;
      }
    }

    if (codePoint <= 0x7f) {
      bytes.push(codePoint);
    } else if (codePoint <= 0x7ff) {
      bytes.push(0xc0 | (codePoint >> 6), 0x80 | (codePoint & 0x3f));
    } else if (codePoint <= 0xffff) {
      bytes.push(
        0xe0 | (codePoint >> 12),
        0x80 | ((codePoint >> 6) & 0x3f),
        0x80 | (codePoint & 0x3f),
      );
    } else {
      bytes.push(
        0xf0 | (codePoint >> 18),
        0x80 | ((codePoint >> 12) & 0x3f),
        0x80 | ((codePoint >> 6) & 0x3f),
        0x80 | (codePoint & 0x3f),
      );
    }
  }
  return bytes;
}

function base64Url(bytes: number[]): string {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let out = "";
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i];
    const b = bytes[i + 1];
    const c = bytes[i + 2];
    out += alphabet[a >> 2];
    out += alphabet[((a & 0x03) << 4) | ((b ?? 0) >> 4)];
    out += b === undefined ? "=" : alphabet[((b & 0x0f) << 2) | ((c ?? 0) >> 6)];
    out += c === undefined ? "=" : alphabet[c & 0x3f];
  }
  return out.replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function shareIdFromUrl(sharingUrl: string): string {
  return `u!${base64Url(encodeUtf8(sharingUrl))}`;
}

function isHttpUrl(value: string): boolean {
  try {
    const parsed = new URL(value);
    return parsed.protocol === "https:" || parsed.protocol === "http:";
  } catch {
    return false;
  }
}

/**
 * @effect readOnly
 * @idempotent
 */
export default async function tool(
  sharing_url: string,
): Promise<unknown> {
  if (typeof sharing_url !== "string" || sharing_url.trim() === "") {
    throw new Error("sharing_url is required");
  }
  const trimmedUrl = sharing_url.trim();
  if (!isHttpUrl(trimmedUrl)) {
    throw new Error("sharing_url must be an http or https URL");
  }

  const share_id = shareIdFromUrl(trimmedUrl);
  const item = await graphGet<GraphDriveItem>(`/shares/${encodeURIComponent(share_id)}/driveItem`);
  const extension = item.name?.split(".").pop()?.toLowerCase() ?? "";
  const drive_id = item.parentReference?.driveId;

  return {
    share_id,
    drive_id,
    item_id: item.id,
    name: item.name,
    webUrl: item.webUrl,
    size: item.size,
    lastModifiedDateTime: item.lastModifiedDateTime,
    file: item.file,
    parentReference: item.parentReference,
    graph_excel_supported: extension === "xlsx",
  };
}
