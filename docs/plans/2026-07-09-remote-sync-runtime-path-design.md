# Remote Sync Runtime PATH Design

## Problem

`subminer sync <host>` can locate a remote launcher while still failing to run
it. The generated launcher uses `#!/usr/bin/env bun`, and non-interactive SSH
shells commonly omit Bun's install directory from `PATH`.

## Design

Remote launcher probes and invocations will use a deterministic `PATH` prefix
covering SubMiner and Bun's supported/common locations:

- `$HOME/.local/bin`
- `$HOME/.bun/bin`
- `/opt/homebrew/bin`
- `/usr/local/bin`
- `/usr/bin`
- `/bin`

Resolution will run `<candidate> --help` under that PATH instead of only using
`command -v`. This confirms both the launcher and its shebang runtime work
before sync creates or transfers snapshots. The same prefixed command returned
by resolution will be used for remote snapshot and merge operations.

`--remote-cmd` remains a launcher path override and receives the same runtime
PATH. Remote commands continue to be shell-quoted where user-controlled.

## Validation

Add focused resolver coverage that simulates an SSH shell where only the
prefixed PATH can start SubMiner. Run launcher unit tests, type checking, docs
tests/build, and the SubMiner launcher/docs verification lanes.
