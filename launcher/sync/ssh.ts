// Re-exported from the shared stats-sync engine so the launcher and the
// app's --sync-cli mode resolve remotes and shell out identically.
export {
  assertSafeSshHost,
  detectRemoteShellFlavor,
  quoteForRemoteShell,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
  shellQuote,
  type RemoteRunResult,
  type RemoteShellFlavor,
} from '../../src/core/services/stats-sync/ssh.js';
