# Sync Stale Socket Guard Design

## Problem

The sync quiescence guard treats any configured mpv socket path as a running
session when the filesystem entry exists. Unix socket files can remain after
mpv exits, so a stale `/tmp/subminer-socket` blocks the remote merge even when
nothing is listening.

## Design

Reuse the launcher's existing Unix socket connection check. The sync guard will
block only when it can connect to the configured socket. Missing, stale, or
otherwise unreachable socket paths will not block sync, and sync will not
delete them.

Because socket connection checks are asynchronous, the sync dispatch, merge
mode, host mode, and quiescence checks will become async. The already-async
launcher entrypoint will await sync dispatch. Snapshot, transfer, merge,
cleanup, and error behavior otherwise remain unchanged.

## Validation

Add focused tests proving a stale socket is ignored and a live socket remains a
hard failure. Re-run sync command and socket tests, type checking, launcher
tests, docs checks, and the SubMiner launcher/docs verification lanes.
