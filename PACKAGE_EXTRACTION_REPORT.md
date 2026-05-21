# Package Extraction Report

Status: repo-ready, not published.

This directory is a clean standalone source tree for:

```text
github.com/solidarity-ai/microsoft-365-toolbox
```

## Included

- `toolbox.devpkg.json`
- `package.json`
- `package-lock.json`
- `tsconfig.json`
- `README.md`
- `PUBLISHING.md`
- `.github/workflows/release.yml`
- `lib/*.ts`
- `tools/*.ts`

## Excluded

- `node_modules/`
- `evals/`
- `eval-runs/`
- `dist/` from source control, except rebuilt local release artifacts
- workshop docs, sessions, handoff files, and consumer workspaces
- stale rebuild planning notes
- the legacy device-code auth helper that can print token JSON

## Preserved Package Identity

- module: `github.com/solidarity-ai/microsoft-365-toolbox`
- package name: `microsoft-365`
- runtime: `typescript-sandbox`
- credential: `microsoft_graph`
- allowed host: `graph.microsoft.com`

## Required Excel Smoke Tools

- `excelSessions.create`
- `excelWorkbooks.inventory`
- `excelRanges.get`
- `excelRanges.update`
- `excelSessions.close`

These tools are present in `toolbox.devpkg.json` and in the rebuilt external
`dist/toolbox.pkg.json`.

## Rebuilt Artifacts

```text
dist/microsoft-365.toolbox.pkg
dist/toolbox.pkg.json
```

Hashes:

```text
3f5f8fcccff29654beb95f731905e4099b07d990953304fbba321be711ac4101  dist/microsoft-365.toolbox.pkg
5f19acc360fdd55780d57ec55e7023b91d6f37cdd057474be0ad8244dc8178c1  dist/toolbox.pkg.json
```

The archive hash matches the `sha256` field in `dist/toolbox.pkg.json`.

## Validation Run

From this directory:

```sh
npm install
npm run typecheck
/opt/homebrew/bin/toolbox info . --json
```

All passed locally.

## Publish Gate

Publishing has not been performed.

To make this independently resolvable for Toolbox, create a GitHub repository:

```text
solidarity-ai/microsoft-365-toolbox
```

Then publish a release tag, for example `v0.1.0`, with these assets attached:

- `microsoft-365.toolbox.pkg`
- `toolbox.pkg.json`

After publication, update the consumer toolset from `v0.0.0` to the real release
version and run non-mutating package/runtime verification.
