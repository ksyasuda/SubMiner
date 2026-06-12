import test from 'node:test';
import assert from 'node:assert/strict';
import { detectWindowsMpvPluginRemovalCandidates } from './first-run-setup-plugin';

test('Windows plugin removal candidates include portable directory installs', () => {
  const mpvPath = 'C:\\tools\\mpv\\mpv.exe';
  const portablePluginDir = 'C:\\tools\\mpv\\portable_config\\scripts\\subminer';
  const existing = new Set([portablePluginDir]);

  const candidates = detectWindowsMpvPluginRemovalCandidates({
    homeDir: 'C:\\Users\\tester',
    appDataDir: 'C:\\Users\\tester\\AppData\\Roaming',
    mpvExecutablePath: mpvPath,
    existsSync: (candidate) => existing.has(candidate),
  });

  assert.deepEqual(candidates, [{ path: portablePluginDir, kind: 'directory' }]);
});
