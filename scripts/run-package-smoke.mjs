import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const resources = process.argv[2];
if (!resources) throw new Error('Usage: bun run test:package <resources-directory>');
const require = createRequire(import.meta.url);
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-package-smoke-'));
const env = { ...process.env, SUBMINER_PACKAGE_SMOKE_DATA: profile };
delete env.ELECTRON_RUN_AS_NODE;
try {
  const result = spawnSync(
    require('electron'),
    [fileURLToPath(new URL('./smoke-package.cjs', import.meta.url)), path.resolve(resources)],
    { env, stdio: 'inherit', timeout: 75_000 },
  );
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} finally {
  fs.rmSync(profile, { recursive: true, force: true, maxRetries: 3 });
}
