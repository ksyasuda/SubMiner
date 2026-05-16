import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import {
  ensureExtensionCopy,
  ensureExtensionCopyAsync,
  shouldCopyYomitanExtension,
} from './yomitan-extension-copy';
import { withSuppressedYomitanExtensionWarnings } from './yomitan-extension-loader';

function makeTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

function writeFile(filePath: string, content: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, 'utf-8');
}

test('suppresses Yomitan contextMenus extension load warnings only while loading', async () => {
  const emitted: string[] = [];
  const warningProcess = {
    emitWarning: (warning: string | Error, options?: { type?: string }) => {
      const message = warning instanceof Error ? warning.message : warning;
      emitted.push(`${options?.type ?? ''}:${message}`);
    },
  } as Pick<NodeJS.Process, 'emitWarning'>;

  await withSuppressedYomitanExtensionWarnings(async () => {
    warningProcess.emitWarning(
      "Warnings loading extension:\nPermission 'contextMenus' is unknown.",
      {
        type: 'ExtensionLoadWarning',
      },
    );
    warningProcess.emitWarning('Other extension warning', { type: 'ExtensionLoadWarning' });
    return null;
  }, warningProcess);

  warningProcess.emitWarning("Permission 'contextMenus' is unknown.", {
    type: 'ExtensionLoadWarning',
  });

  assert.deepEqual(emitted, [
    'ExtensionLoadWarning:Other extension warning',
    "ExtensionLoadWarning:Permission 'contextMenus' is unknown.",
  ]);
});

test('shouldCopyYomitanExtension detects popup runtime script drift', () => {
  const tempRoot = makeTempDir('subminer-yomitan-copy-');
  const sourceDir = path.join(tempRoot, 'source');
  const targetDir = path.join(tempRoot, 'target');

  writeFile(path.join(sourceDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  writeFile(path.join(targetDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  writeFile(path.join(sourceDir, 'js', 'app', 'popup.js'), 'same-popup-script');
  writeFile(path.join(targetDir, 'js', 'app', 'popup.js'), 'same-popup-script');
  writeFile(path.join(sourceDir, 'js', 'display', 'popup-main.js'), 'source-popup-main');
  writeFile(path.join(targetDir, 'js', 'display', 'popup-main.js'), 'target-popup-main');

  assert.equal(shouldCopyYomitanExtension(sourceDir, targetDir), true);
});

test('shouldCopyYomitanExtension skips copy when extension contents match', () => {
  const tempRoot = makeTempDir('subminer-yomitan-copy-');
  const sourceDir = path.join(tempRoot, 'source');
  const targetDir = path.join(tempRoot, 'target');

  writeFile(path.join(sourceDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  writeFile(path.join(targetDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  writeFile(path.join(sourceDir, 'js', 'app', 'popup.js'), 'same-popup-script');
  writeFile(path.join(targetDir, 'js', 'app', 'popup.js'), 'same-popup-script');
  writeFile(path.join(sourceDir, 'js', 'display', 'popup-main.js'), 'same-popup-main');
  writeFile(path.join(targetDir, 'js', 'display', 'popup-main.js'), 'same-popup-main');
  writeFile(path.join(sourceDir, 'js', 'display', 'display.js'), 'same-display');
  writeFile(path.join(targetDir, 'js', 'display', 'display.js'), 'same-display');
  writeFile(path.join(sourceDir, 'js', 'display', 'display-audio.js'), 'same-display-audio');
  writeFile(path.join(targetDir, 'js', 'display', 'display-audio.js'), 'same-display-audio');

  assert.equal(shouldCopyYomitanExtension(sourceDir, targetDir), false);
});

test('ensureExtensionCopy refreshes copied extension when display files change', () => {
  const sourceRoot = makeTempDir('subminer-yomitan-src-');
  const userDataRoot = makeTempDir('subminer-yomitan-user-');

  const sourceDir = path.join(sourceRoot, 'yomitan');
  const targetDir = path.join(userDataRoot, 'extensions', 'yomitan');

  fs.mkdirSync(path.join(sourceDir, 'js', 'display'), { recursive: true });
  fs.mkdirSync(path.join(targetDir, 'js', 'display'), { recursive: true });

  fs.writeFileSync(path.join(sourceDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  fs.writeFileSync(path.join(targetDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  fs.writeFileSync(
    path.join(sourceDir, 'js', 'display', 'structured-content-generator.js'),
    'new display code',
  );
  fs.writeFileSync(
    path.join(targetDir, 'js', 'display', 'structured-content-generator.js'),
    'old display code',
  );

  const result = ensureExtensionCopy(sourceDir, userDataRoot);

  assert.equal(result.targetDir, targetDir);
  assert.equal(result.copied, true);
  assert.equal(
    fs.readFileSync(
      path.join(targetDir, 'js', 'display', 'structured-content-generator.js'),
      'utf8',
    ),
    'new display code',
  );
});

test('ensureExtensionCopyAsync refreshes copied extension without completing synchronously', async () => {
  const sourceRoot = makeTempDir('subminer-yomitan-src-');
  const userDataRoot = makeTempDir('subminer-yomitan-user-');

  const sourceDir = path.join(sourceRoot, 'yomitan');
  const targetDir = path.join(userDataRoot, 'extensions', 'yomitan');

  fs.mkdirSync(path.join(sourceDir, 'js', 'display'), { recursive: true });
  fs.mkdirSync(path.join(targetDir, 'js', 'display'), { recursive: true });

  fs.writeFileSync(path.join(sourceDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  fs.writeFileSync(path.join(targetDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  fs.writeFileSync(
    path.join(sourceDir, 'js', 'display', 'structured-content-generator.js'),
    'new display code',
  );
  fs.writeFileSync(
    path.join(targetDir, 'js', 'display', 'structured-content-generator.js'),
    'old display code',
  );

  let completed = false;
  const resultPromise = ensureExtensionCopyAsync(sourceDir, userDataRoot).then((result) => {
    completed = true;
    return result;
  });

  assert.equal(completed, false);

  const result = await resultPromise;

  assert.equal(result.targetDir, targetDir);
  assert.equal(result.copied, true);
  assert.equal(
    fs.readFileSync(
      path.join(targetDir, 'js', 'display', 'structured-content-generator.js'),
      'utf8',
    ),
    'new display code',
  );
});

test('ensureExtensionCopyAsync shares an in-flight refresh for the same copied extension', async () => {
  const sourceRoot = makeTempDir('subminer-yomitan-src-');
  const userDataRoot = makeTempDir('subminer-yomitan-user-');

  const sourceDir = path.join(sourceRoot, 'yomitan');
  const targetDir = path.join(userDataRoot, 'extensions', 'yomitan');

  fs.mkdirSync(path.join(sourceDir, 'js', 'pages', 'settings'), { recursive: true });
  fs.mkdirSync(path.join(targetDir, 'js', 'pages', 'settings'), { recursive: true });

  fs.writeFileSync(path.join(sourceDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  fs.writeFileSync(path.join(targetDir, 'manifest.json'), JSON.stringify({ version: '1.0.0' }));
  fs.writeFileSync(
    path.join(sourceDir, 'js', 'pages', 'settings', 'settings-main.js'),
    'new settings code',
  );
  fs.writeFileSync(
    path.join(targetDir, 'js', 'pages', 'settings', 'settings-main.js'),
    'old settings code',
  );

  const originalCp = fs.promises.cp;
  let cpCalls = 0;
  let firstCopyStarted = false;
  let releaseFirstCopy: () => void = () => {};
  const firstCopyStartedPromise = new Promise<void>((resolve) => {
    const fsPromises = fs.promises as typeof fs.promises & {
      cp: typeof fs.promises.cp;
    };
    fsPromises.cp = async (...args: Parameters<typeof fs.promises.cp>) => {
      cpCalls++;
      if (!firstCopyStarted) {
        firstCopyStarted = true;
        resolve();
        await new Promise<void>((release) => {
          releaseFirstCopy = release;
        });
      }
      return await originalCp(...args);
    };
  });

  try {
    const first = ensureExtensionCopyAsync(sourceDir, userDataRoot);
    await firstCopyStartedPromise;
    const second = ensureExtensionCopyAsync(sourceDir, userDataRoot);

    releaseFirstCopy();
    const results = await Promise.all([first, second]);

    assert.equal(cpCalls, 1);
    assert.equal(results[0].targetDir, targetDir);
    assert.equal(results[1].targetDir, targetDir);
    assert.equal(results[0].copied, true);
    assert.equal(results[1].copied, true);
    assert.equal(
      fs.readFileSync(path.join(targetDir, 'js', 'pages', 'settings', 'settings-main.js'), 'utf8'),
      'new settings code',
    );
  } finally {
    const fsPromises = fs.promises as typeof fs.promises & {
      cp: typeof fs.promises.cp;
    };
    fsPromises.cp = originalCp;
  }
});
