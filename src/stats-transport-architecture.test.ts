import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

const root = process.cwd();

function read(relativePath: string): string {
  return fs.readFileSync(path.join(root, relativePath), 'utf8');
}

test('stats data uses the shared HTTP contract while native dialogs retain IPC', () => {
  assert.equal(fs.existsSync(path.join(root, 'stats/src/lib/ipc-client.ts')), false);

  const preload = read('src/preload-stats.ts');
  assert.doesNotMatch(preload, /statsGet[A-Z]/);
  assert.match(preload, /statsNativeConfirmDialog/);

  const ipcContracts = read('src/shared/ipc/contracts.ts');
  assert.doesNotMatch(ipcContracts, /statsGet[A-Z]/);

  const ipcHandlers = read('src/core/services/ipc.ts');
  assert.doesNotMatch(ipcHandlers, /statsGet[A-Z]/);

  const apiClient = read('stats/src/lib/api-client.ts');
  const statsServerRoutes = [
    'analytics-routes.ts',
    'integration-routes.ts',
    'library-routes.ts',
    'mining-routes.ts',
  ].map((filename) => ({
    filename,
    source: read(`src/core/services/stats-server/${filename}`),
  }));
  assert.match(apiClient, /StatsHttpClient/);
  for (const { filename, source } of statsServerRoutes) {
    assert.match(source, /statsJson/, `${filename} must use the shared HTTP contract`);
  }
});
