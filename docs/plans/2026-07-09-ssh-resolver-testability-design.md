# SSH Resolver Testability Design

## Problem

The remote launcher resolver test installs a temporary `ssh` executable and
mutates `process.env.PATH` after Bun starts. Bun's synchronous process launcher
does not reliably use that late PATH change, so the focused test fails locally
and in CI even though the command-building behavior is correct.

## Design

Allow `resolveRemoteSubminerCommand` to accept an optional remote runner that
defaults to the production `runSsh` function. Existing callers retain the same
behavior and two-argument API. The focused test will provide a small runner,
exercise the real candidate and PATH construction, and assert the host and
remote command passed across the SSH boundary.

Alternatives rejected:

- Explicitly forwarding `process.env` to `spawnSync` still depends on Bun's
  executable-resolution behavior and changes production code to support a test
  harness.
- Module-level mocking couples the test to loader behavior and obscures the
  resolver's actual dependency.

## Validation

First update the focused test and confirm it remains red while the resolver
ignores the injected runner. Then add the minimal runner parameter and confirm
the focused test passes. Run the launcher verification lane, type checking,
and the repository's default handoff gate.
