import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { enforceElectronRuntimeGuard, SUPPORTED_ELECTRON_MAJOR } from './electron-runtime-guard';

function withTempDir(run: (directory: string) => void): void {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-electron-guard-'));
  try {
    run(directory);
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
}

test('runtime guard major matches the pinned Electron dependency', () => {
  const packageJson = JSON.parse(
    fs.readFileSync(path.join(process.cwd(), 'package.json'), 'utf8'),
  ) as { devDependencies: { electron: string } };

  assert.equal(Number.parseInt(packageJson.devDependencies.electron.split('.', 1)[0]!, 10), 43);
  assert.equal(SUPPORTED_ELECTRON_MAJOR, 43);
});

test('runtime guard records the supported Electron major', () => {
  withTempDir((userDataPath) => {
    const result = enforceElectronRuntimeGuard({
      electronVersion: '43.4.1',
      userDataPath,
      supportedElectronMajor: 43,
    });

    assert.equal(result.ok, true);
    assert.deepEqual(JSON.parse(fs.readFileSync(result.statePath, 'utf8')), {
      highestElectronMajor: 43,
      lastElectronVersion: '43.4.1',
    });
  });
});

test('runtime guard rejects a runtime outside the build major without writing state', () => {
  withTempDir((userDataPath) => {
    const result = enforceElectronRuntimeGuard({
      electronVersion: '44.0.0',
      userDataPath,
      supportedElectronMajor: 43,
    });

    assert.equal(result.ok, false);
    if (result.ok) return;
    assert.equal(result.title, 'Unsupported Electron runtime');
    assert.match(result.details, /requires Electron 43/);
    assert.equal(fs.existsSync(result.statePath), false);
  });
});

test('runtime guard rejects prerelease Electron versions without writing state', () => {
  withTempDir((userDataPath) => {
    const result = enforceElectronRuntimeGuard({
      electronVersion: '43.4.1-beta.1',
      userDataPath,
      supportedElectronMajor: 43,
    });

    assert.equal(result.ok, false);
    if (result.ok) return;
    assert.equal(result.title, 'SubMiner could not verify Electron');
    assert.equal(fs.existsSync(result.statePath), false);
  });
});

test('runtime guard blocks a profile downgrade before rewriting its safety record', () => {
  withTempDir((userDataPath) => {
    const statePath = path.join(userDataPath, 'electron-runtime.json');
    fs.writeFileSync(
      statePath,
      JSON.stringify({ highestElectronMajor: 44, lastElectronVersion: '44.1.0' }),
      'utf8',
    );

    const result = enforceElectronRuntimeGuard({
      electronVersion: '43.4.1',
      userDataPath,
      supportedElectronMajor: 43,
    });

    assert.equal(result.ok, false);
    if (result.ok) return;
    assert.equal(result.title, 'Electron downgrade blocked');
    assert.match(result.details, /destroy Yomitan dictionaries/);
    assert.deepEqual(JSON.parse(fs.readFileSync(statePath, 'utf8')), {
      highestElectronMajor: 44,
      lastElectronVersion: '44.1.0',
    });
  });
});

test('runtime guard blocks a downgrade within the supported Electron major', () => {
  withTempDir((userDataPath) => {
    const statePath = path.join(userDataPath, 'electron-runtime.json');
    fs.writeFileSync(
      statePath,
      JSON.stringify({ highestElectronMajor: 43, lastElectronVersion: '43.4.1' }),
      'utf8',
    );

    const result = enforceElectronRuntimeGuard({
      electronVersion: '43.3.0',
      userDataPath,
      supportedElectronMajor: 43,
    });

    assert.equal(result.ok, false);
    if (result.ok) return;
    assert.equal(result.title, 'Electron downgrade blocked');
    assert.deepEqual(JSON.parse(fs.readFileSync(statePath, 'utf8')), {
      highestElectronMajor: 43,
      lastElectronVersion: '43.4.1',
    });
  });
});

test('runtime guard fails closed when its safety record is malformed', () => {
  withTempDir((userDataPath) => {
    const statePath = path.join(userDataPath, 'electron-runtime.json');
    fs.writeFileSync(statePath, '{}', 'utf8');

    const result = enforceElectronRuntimeGuard({
      electronVersion: '43.4.1',
      userDataPath,
      supportedElectronMajor: 43,
    });

    assert.equal(result.ok, false);
    if (result.ok) return;
    assert.equal(result.title, 'SubMiner profile safety check failed');
    assert.match(result.details, /invalid format/);
  });
});
