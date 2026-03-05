import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { shouldCopyYomitanExtension } from './yomitan-extension-copy';

function writeFile(filePath: string, content: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, 'utf-8');
}

test('shouldCopyYomitanExtension detects popup runtime script drift', () => {
  const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'yomitan-copy-test-'));
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

test('shouldCopyYomitanExtension skips copy when versions and watched scripts match', () => {
  const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'yomitan-copy-test-'));
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
