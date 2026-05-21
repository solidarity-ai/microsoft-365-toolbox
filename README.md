# Microsoft 365 Toolbox Package

Preview Toolbox package for Microsoft Graph and Excel workbook tools.

This package targets the installed preview `toolbox` auth model. Tools do not accept access tokens and do not set `Authorization`; the package declares a Microsoft Graph OAuth2 credential in `toolbox.devpkg.json`, and Toolbox injects bearer auth into allowed `graph.microsoft.com` requests.

## Tools

The package includes the original read-only Microsoft 365 discovery tools plus a V1 Excel surface for workbook discovery, sessions, inventory, range reads, scoped session updates, calculation, formula/error scanning, and diff checks.

| Tool file | Tool name | Graph call | Delegated scope |
| --- | --- | --- | --- |
| `tools/me.get.ts` | `me.get` | `GET /me` | `User.Read` |
| `tools/sites.search.ts` | `sites.search` | `GET /sites?search={query}` | `Sites.Read.All` |
| `tools/sites.get.ts` | `sites.get` | `GET /sites/{site_id}` | `Sites.Read.All` |
| `tools/site-lists.list.ts` | `siteLists.list` | `GET /sites/{site_id}/lists` | `Sites.Read.All` |
| `tools/site-list-items.list.ts` | `siteListItems.list` | `GET /sites/{site_id}/lists/{list_id}/items?$expand=fields` | `Sites.Read.All` |
| `tools/site-drives.list.ts` | `siteDrives.list` | `GET /sites/{site_id}/drives` | `Files.Read.All` or tenant-approved SharePoint read access |
| `tools/drive-items.get.ts` | `driveItems.get` | `GET /drives/{drive_id}/items/{item_id}` | `Files.Read.All` |

Excel V1 tools:

| Tool file | Tool name | Purpose |
| --- | --- | --- |
| `tools/excel-workbooks.find.ts` | `excelWorkbooks.find` | Find candidate `.xlsx`, `.xlsm`, and `.xlsb` files in OneDrive/SharePoint search. |
| `tools/excel-workbooks.get.ts` | `excelWorkbooks.get` | Return workbook metadata and capability flags. |
| `tools/excel-sessions.create.ts` | `excelSessions.create` | Create non-persistent or persistent Graph workbook sessions. |
| `tools/excel-sessions.close.ts` | `excelSessions.close` | Close a Graph workbook session. |
| `tools/excel-workbooks.inventory.ts` | `excelWorkbooks.inventory` | Summarize worksheets, used ranges, tables, and named items. |
| `tools/excel-worksheets.list.ts` | `excelWorksheets.list` | List workbook worksheets. |
| `tools/excel-ranges.used.ts` | `excelRanges.used` | Read worksheet used-range metadata and optional values. |
| `tools/excel-ranges.get.ts` | `excelRanges.get` | Read a scoped range with selectable values, formulas, text, and formats. |
| `tools/excel-ranges.update.ts` | `excelRanges.update` | Patch values/formulas/formats inside an explicit workbook session. |
| `tools/excel-formulas.scan.ts` | `excelFormulas.scan` | Summarize formula patterns and risky/external formula usage. |
| `tools/excel-workbooks.calculate.ts` | `excelWorkbooks.calculate` | Trigger Graph workbook calculation. |
| `tools/excel-errors.scan.ts` | `excelErrors.scan` | Scan used ranges for Excel error values. |
| `tools/excel-diffs.compare-session-to-workbook.ts` | `excelDiffs.compareSessionToWorkbook` | Compare a post-preview range to caller-supplied baseline values. |

## Microsoft Entra Setup

Create or reuse a Microsoft Entra app registration with delegated Microsoft Graph permissions:

- `openid`
- `profile`
- `offline_access`
- `User.Read`
- `Sites.Read.All`
- `Files.Read.All`
- `Files.ReadWrite.All`

Use a public client app for the preview PKCE flow. When `toolbox auth` prompts for `client_secret`, press Enter to skip it for a public client.

The installed preview Toolbox currently uses the built-in Microsoft provider endpoints under the `common` tenant. If your organization requires a single-tenant authority, this package manifest will need tenant-specific `auth_url` and `token_url` provider values instead of `{ "name": "microsoft" }`.

## Validate Package Metadata

From the package repo:

```sh
/opt/homebrew/bin/toolbox info .
```

Use `--json` when you want to inspect the resolved credential and allowed host metadata.

## Release

The canonical package module is:

```text
github.com/solidarity-ai/microsoft-365-toolbox
```

For GitHub-hosted package resolution, publish a GitHub Release whose tag matches
the toolset version, for example `v0.1.0`. Attach both built artifacts:

- `microsoft-365.toolbox.pkg`
- `toolbox.pkg.json`

Toolbox expects the release assets to contain the archive and the external
manifest with the archive SHA-256.

## Install From Consumer Workspace

A consumer workspace declares a real release version:

```sh
cd /path/to/m365-toolbox-user
/opt/homebrew/bin/toolbox install
```

`toolbox.toolset.json` should declare
`github.com/solidarity-ai/microsoft-365-toolbox` at a resolvable release version.
Local development can replace that module with a local checkout, but that is not
independent package evidence.

## Authenticate

Run auth against the package directory:

```sh
/opt/homebrew/bin/toolbox auth .
```

The command prompts for the Microsoft app `client_id`, optionally a `client_secret`, then opens a browser for OAuth consent. For the PKCE public-client flow, leave `client_secret` blank.

To check stored credential state:

```sh
/opt/homebrew/bin/toolbox auth --check .
```

## Serve MCP

From the external-user workspace:

```sh
cd /path/to/m365-toolbox-user
/opt/homebrew/bin/toolbox mcp
```

Connect your MCP client to that stdio command. Tool schemas should not include an `access_token` parameter; bearer auth is injected by Toolbox when authenticated tools fetch `https://graph.microsoft.com`.

## Development

```sh
npm install
npm run typecheck
```

## License

AGPL-3.0-or-later.
