# Publishing

This package is intended to resolve as:

```text
github.com/solidarity-ai/microsoft-365-toolbox@<version>
```

The release assets must include:

- `microsoft-365.toolbox.pkg`
- `toolbox.pkg.json`

The external `toolbox.pkg.json` must contain the SHA-256 of the archive in its
`sha256` field.

## Build Artifacts Locally

From the tools-workshop checkout that contains the Toolbox preview source:

```sh
GOCACHE=/private/tmp/go-build-cache \
  go run ./cmd/toolbox-pack \
  -out /path/to/microsoft-365-toolbox/dist \
  /path/to/microsoft-365-toolbox
```

The command writes the archive and external manifest into `dist/`.

## GitHub Release

Create a release whose tag is the version used in consumer toolsets, for example
`v0.1.0`, and attach the two files from `dist/`.

After the release is available, verify from a consumer workspace:

```sh
/opt/homebrew/bin/toolbox install -t toolbox.toolset.json
/opt/homebrew/bin/toolbox info github.com/solidarity-ai/microsoft-365-toolbox -t toolbox.toolset.json --json
```

Do not treat a local replacement or source checkout as independent package
evidence.
