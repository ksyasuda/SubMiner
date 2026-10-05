import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { getRunCommand } from './command-line-launcher-deps';
import { windowsLauncherBootstrapContent } from './windows-launcher-bootstrap';

function workspace(t: test.TestContext): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer windows bootstrap '));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  return root;
}

test('configured app path is an escaped fallback after the environment override', () => {
  const configured = 'D:\\Apps & Tools\\SubMiner 100% !\\SubMiner.exe';
  const content = windowsLauncherBootstrapContent(configured);

  assert.ok(
    content.indexOf('if defined SUBMINER_BINARY_PATH goto subminer_app_found') <
      content.indexOf('if not exist "D:\\Apps & Tools\\SubMiner 100%% !\\SubMiner.exe"'),
  );
  assert.match(
    content,
    /set "SUBMINER_BINARY_PATH=D:\\Apps & Tools\\SubMiner 100%% !\\SubMiner\.exe"/,
  );
  assert.throws(
    () => windowsLauncherBootstrapContent('C:\\Bad "Install"\\SubMiner.exe'),
    /quotes or newlines/,
  );
});

// Cold and cached launches each allow 15 seconds, plus executable copies during setup.
test(
  'Windows bootstrap prepares once and forwards metacharacter arguments',
  { timeout: 45_000 },
  async (t) => {
    if (process.platform !== 'win32') return;

    const root = workspace(t);
    const appDirectory = path.join(root, 'Installed & App 100% !');
    const appPath = path.join(appDirectory, 'SubMiner.exe');
    const resourcesPath = path.join(appDirectory, 'resources');
    const launcherDirectory = path.join(resourcesPath, 'launcher');
    const localAppData = path.join(root, 'Local App Data');
    const bootstrapPath = path.join(root, 'subminer.cmd');
    const version = '1.2.3-test';
    const cachedBunPath = path.join(
      localAppData,
      'SubMiner',
      'launcher-runtime',
      version,
      'bun.exe',
    );
    fs.mkdirSync(launcherDirectory, { recursive: true });
    fs.copyFileSync(process.execPath, appPath);
    fs.writeFileSync(path.join(launcherDirectory, 'version'), version);
    fs.writeFileSync(
      path.join(launcherDirectory, 'subminer.js'),
      'console.log(JSON.stringify({args:process.argv.slice(2),app:process.env.SUBMINER_BINARY_PATH,resources:process.env.SUBMINER_RESOURCES_PATH,managed:process.env.SUBMINER_MANAGED_LAUNCHER}));',
    );
    fs.writeFileSync(
      path.join(launcherDirectory, 'prepare.cjs'),
      `const fs=require('node:fs');const path=require('node:path');exports.prepareLauncherRuntime=()=>{const target=${JSON.stringify(cachedBunPath)};fs.mkdirSync(path.dirname(target),{recursive:true});fs.copyFileSync(process.execPath,target);};`,
    );
    fs.writeFileSync(bootstrapPath, windowsLauncherBootstrapContent(appPath));

    const args = [
      'spaces here',
      '100%',
      'p%TEMP%q',
      'bang!',
      'a&b',
      'x|y',
      '<left>',
      'caret^',
      'say "hi"',
      '日本語',
    ];
    const env = { ...process.env, PATH: '', Path: '', LOCALAPPDATA: localAppData };
    const first = await getRunCommand({})(bootstrapPath, args, { env });
    assert.equal(first.exitCode, 0, first.stderr);
    assert.ok(fs.existsSync(cachedBunPath));
    const firstPayload: unknown = JSON.parse(first.stdout);
    assert.deepEqual(firstPayload, {
      args,
      app: appPath,
      resources: resourcesPath,
      managed: '1',
    });

    fs.writeFileSync(
      path.join(launcherDirectory, 'prepare.cjs'),
      "throw new Error('cached launch should not prepare');",
    );
    const second = await getRunCommand({})(bootstrapPath, args, { env });
    assert.equal(second.exitCode, 0, second.stderr);
    const secondPayload: unknown = JSON.parse(second.stdout);
    assert.deepEqual(secondPayload, firstPayload);
  },
);
