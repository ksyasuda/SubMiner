import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';
import { posixLauncherBootstrapContent, shellQuote } from './posix-launcher-bootstrap';

test('downloaded launcher prepares once, survives unmount, and refreshes after app replacement', (t) => {
  if (process.platform !== 'linux') return;
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "subminer bootstrap's "));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const appPath = path.join(root, 'SubMiner.AppImage');
  const resources = path.join(root, 'mount', 'resources');
  fs.mkdirSync(path.join(resources, 'launcher'), { recursive: true });
  fs.mkdirSync(path.join(resources, 'bun', 'licenses'), { recursive: true });
  fs.symlinkSync(process.execPath, path.join(resources, 'bun', 'bun'));
  fs.writeFileSync(path.join(resources, 'bun', 'licenses', 'notice'), 'Bun notice');
  fs.writeFileSync(path.join(resources, 'launcher', 'version'), '1.0.0\n');
  fs.writeFileSync(
    path.join(resources, 'launcher', 'prepare.cjs'),
    `exports.prepareLauncherRuntime = require(${JSON.stringify(path.join(__dirname, 'prepare-launcher-runtime.ts'))}).prepareLauncherRuntime;`,
  );
  const script = (version: number) =>
    `console.log(JSON.stringify({version:${version},args:process.argv.slice(2)}));`;
  fs.writeFileSync(path.join(resources, 'launcher', 'subminer.js'), script(1));
  const prepares = path.join(root, 'preparations');
  const appContent = `#!/bin/sh\necho prepared >> ${shellQuote(prepares)}\nexport APPDIR=${shellQuote(path.dirname(resources))}\nexec ${shellQuote(process.execPath)} "$@"\n`;
  fs.writeFileSync(appPath, appContent, { mode: 0o755 });
  const wrapper = path.join(root, 'subminer');
  fs.writeFileSync(wrapper, posixLauncherBootstrapContent(), { mode: 0o755 });
  const env = { HOME: root, PATH: '', SUBMINER_BINARY_PATH: appPath };
  const args = ['space here', "a'b", 'a&b', '$(touch nope)', '日本語'];
  const run = (extraEnv = {}) =>
    spawnSync(wrapper, args, { env: { ...env, ...extraEnv }, encoding: 'utf8' });
  const first = run();
  assert.equal(first.status, 0, first.stderr);
  assert.deepEqual(JSON.parse(first.stdout), { version: 1, args });
  const cache = path.join(root, '.local', 'share', 'SubMiner', 'launcher');
  assert.equal(fs.readFileSync(path.join(cache, 'licenses', 'notice'), 'utf8'), 'Bun notice');

  fs.renameSync(resources, `${resources}.unmounted`);
  const warm = run({ SUBMINER_BINARY_PATH: '' }); // Finds the app recorded by preparation.
  assert.equal(warm.status, 0, warm.stderr);
  assert.deepEqual(JSON.parse(warm.stdout), { version: 1, args });
  assert.equal(fs.readFileSync(prepares, 'utf8'), 'prepared\n');
  fs.renameSync(`${resources}.unmounted`, resources);

  // Same version and path, new inode: manual replacement must still refresh.
  fs.writeFileSync(`${appPath}.new`, appContent, { mode: 0o755 });
  fs.renameSync(`${appPath}.new`, appPath);
  fs.writeFileSync(path.join(resources, 'launcher', 'subminer.js'), script(2));
  const updated = run();
  assert.equal(updated.status, 0, updated.stderr);
  assert.deepEqual(JSON.parse(updated.stdout), { version: 2, args });
  assert.equal(fs.readFileSync(prepares, 'utf8'), 'prepared\nprepared\n');

  fs.unlinkSync(path.join(cache, 'bun'));
  assert.equal(run().status, 0);
  assert.equal(fs.readFileSync(prepares, 'utf8'), 'prepared\nprepared\nprepared\n');

  fs.writeFileSync(appPath, '#!/bin/sh\nexit 42\n', { mode: 0o755 });
  const failed = run();
  assert.equal(failed.status, 42);
  assert.equal(failed.stdout, ''); // Never silently execute stale CLI after failed preparation.
});
