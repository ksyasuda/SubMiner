/**
 * Electron AppImages export `LD_LIBRARY_PATH=<mount>/usr/lib`, whose bundled libraries (libnotify and
 * friends) predate the symbols system binaries link against, so an inherited environment can kill a
 * spawned child with a symbol lookup error. A system binary resolves its own libraries fine, so the
 * override is dropped entirely rather than filtered.
 */
export function buildSystemBinaryEnv(env: NodeJS.ProcessEnv = process.env): NodeJS.ProcessEnv {
  const { LD_LIBRARY_PATH: _dropped, ...rest } = env;
  return rest;
}
