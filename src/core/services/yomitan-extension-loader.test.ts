import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { ensureExtensionCopy, shouldCopyYomitanExtension } from './yomitan-extension-copy';

function makeTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

function writeFile(filePath: string, content: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, 'utf-8');
}

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
