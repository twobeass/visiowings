# Configuration

Most users only ever touch two flags: `--file` and `--bidirectional`.
For everything else, persist your defaults in a `.visiowings.toml` so
you can run `visiowings edit` without arguments.

## Quick start with `visiowings init`

```bash
cd path/to/your/project
visiowings init
```

The wizard:

1. Lists the Visio documents currently open via COM.
2. Lets you pick the main document (or enter a path manually).
3. Asks for the output directory, bidirectional sync, Rubberduck,
   and codepage preferences.
4. Writes a `.visiowings.toml` next to your project.

Subsequent commands inherit these defaults:

```bash
visiowings edit          # implies --file from config
visiowings export -o vba/2024  # explicit args still win
```

## File format

```toml
# .visiowings.toml
file = "drawings/main.vsdm"
output = "vba"
codepage = "cp1252"            # blank/missing = auto-detect from doc
bidirectional = true
rubberduck = false
sync_delete_modules = false
force = false
```

## Discovery rules

`visiowings` walks up the directory tree from the working directory
looking for the first `.visiowings.toml` it can find — same convention
as `pyproject.toml` and `.git/`. Override with the explicit flag if you
need to:

```bash
visiowings edit --file other.vsdx
```

## Disable the update check

The CLI fires a daily background HEAD against PyPI to check for new
releases. Disable it any of three ways:

```bash
export VISIOWINGS_NO_UPDATE_CHECK=1   # session-wide
visiowings edit --no-update-check     # per-invocation
# or run in CI; the CI env var disables it automatically
```
