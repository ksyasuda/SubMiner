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

test('generic bootstrap searches supported Windows install locations in order', () => {
  const content = windowsLauncherBootstrapContent();
  const override = content.indexOf('if defined SUBMINER_BINARY_PATH goto subminer_app_found');
  const localInstall = content.indexOf('%LOCALAPPDATA%\\Programs\\SubMiner\\SubMiner.exe');
  const machineInstall = content.indexOf('%ProgramFiles%\\SubMiner\\SubMiner.exe');

  assert.ok(override >= 0);
  assert.ok(localInstall > override);
  assert.ok(machineInstall > localInstall);
  assert.match(content, /SubMiner app not found/);
});

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

test('cached runtime is the fast path and launcher arguments remain opaque to the batch file', () => {
  const content = windowsLauncherBootstrapContent();
  const cacheCheck = content.indexOf('if exist "%SUBMINER_BUN_PATH%" goto subminer_run');
  const electronPrepare = content.indexOf('set "ELECTRON_RUN_AS_NODE=1"');
  const run = content.indexOf(
    '"%SUBMINER_BUN_PATH%" "%SUBMINER_RESOURCES_PATH%\\launcher\\subminer.js" %*',
  );

  assert.match(content, /^@echo off\r\nrem SubMiner managed launcher \(bundled runtime\)/);
  assert.match(content, /setlocal DisableDelayedExpansion/);
  assert.ok(cacheCheck >= 0);
  assert.ok(electronPrepare > cacheCheck);
  assert.ok(run > electronPrepare);
  assert.doesNotMatch(content, /powershell/i);
  assert.doesNotMatch(content, /^\s*(?:call\s+)?bun(?:\.exe)?(?:\s|$)/im);
});

test('preparation failures keep their exit code and never continue to the launcher', () => {
  const content = windowsLauncherBootstrapContent();

  assert.match(content, /set "SUBMINER_PREPARE_EXIT=%errorlevel%"/);
  assert.match(content, /if not "%SUBMINER_PREPARE_EXIT%"=="0" exit \/b %SUBMINER_PREPARE_EXIT%/);
  assert.match(content, /if not exist "%SUBMINER_BUN_PATH%" goto subminer_prepare_missing/);
  assert.match(content, /set "ELECTRON_RUN_AS_NODE="/);
});

test('Windows bootstrap prepares once and forwards metacharacter arguments', async (t) => {
  if (process.platform !== 'win32') return;

  const root = workspace(t);
  const appDirectory = path.join(root, 'Installed & App 100% !');
  const appPath = path.join(appDirectory, 'SubMiner.exe');
  const resourcesPath = path.join(appDirectory, 'resources');
  const launcherDirectory = path.join(resourcesPath, 'launcher');
  const localAppData = path.join(root, 'Local App Data');
  const bootstrapPath = path.join(root, 'subminer.cmd');
  const version = '1.2.3-test';
  const cachedBunPath = path.join(localAppData, 'SubMiner', 'launcher-runtime', version, 'bun.exe');
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
});
