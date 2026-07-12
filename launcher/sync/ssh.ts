// Re-exported from the shared stats-sync engine so the launcher and the
// app's --sync-cli mode resolve remotes and shell out identically.
export {
  assertSafeSshHost,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
  shellQuote,
  type RemoteRunResult,
} from '../../src/core/services/stats-sync/ssh.js';
