import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import {
  findRofiTheme,
  findRofiThumbnailerDataRoot,
  formatRofiPrompt,
  prependXdgDataDir,
} from './picker';

// ── formatRofiPrompt: spacing between prompt and input field ──────────────────

test('formatRofiPrompt appends a single trailing space', () => {
  assert.equal(formatRofiPrompt('Select Video'), 'Select Video ');
});

test('formatRofiPrompt collapses existing trailing whitespace to one space', () => {
  assert.equal(formatRofiPrompt('Watch History   '), 'Watch History ');
});

test('formatRofiPrompt leaves an empty prompt empty', () => {
  assert.equal(formatRofiPrompt(''), '');
  assert.equal(formatRofiPrompt('   '), '');
});

// ── findRofiTheme: Linux packaged path discovery ──────────────────────────────

const ROFI_THEME_FILE = 'subminer.rasi';
const ROFI_THUMBNAILER_FILE = 'subminer-ffmpegthumbnailer.thumbnailer';

function makeFile(filePath: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, '/* theme */');
}

function withPlatform<T>(platform: NodeJS.Platform, callback: () => T): T {
  const originalDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    value: platform,
  });
  try {
    return callback();
  } finally {
    if (originalDescriptor) {
      Object.defineProperty(process, 'platform', originalDescriptor);
    }
  }
}

test('findRofiTheme resolves /usr/local/share/SubMiner/themes/subminer.rasi when it exists', () => {
  const originalExistsSync = fs.existsSync;
  const targetPath = `/usr/local/share/SubMiner/themes/${ROFI_THEME_FILE}`;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).existsSync = (filePath: unknown): boolean => {
      if (filePath === targetPath) return true;
      return false;
    };

    const result = withPlatform('linux', () => findRofiTheme('/usr/local/bin/subminer'));
    assert.equal(result, targetPath);
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).existsSync = originalExistsSync;
  }
});

test('findRofiTheme resolves /usr/share/SubMiner/themes/subminer.rasi when /usr/local/share one does not exist', () => {
  const originalExistsSync = fs.existsSync;
  const localSharePath = `/usr/local/share/SubMiner/themes/${ROFI_THEME_FILE}`;
  const sharePath = `/usr/share/SubMiner/themes/${ROFI_THEME_FILE}`;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).existsSync = (filePath: unknown): boolean => {
      if (filePath === sharePath) return true;
      if (filePath === localSharePath) return false;
      return false;
    };

    const result = withPlatform('linux', () => findRofiTheme('/usr/bin/subminer'));
    assert.equal(result, sharePath);
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).existsSync = originalExistsSync;
  }
});

test('findRofiTheme resolves XDG_DATA_HOME/SubMiner/themes/subminer.rasi when set and file exists', () => {
  const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-xdg-'));
  const originalXdgDataHome = process.env.XDG_DATA_HOME;
  try {
    process.env.XDG_DATA_HOME = baseDir;
    const themePath = path.join(baseDir, `SubMiner/themes/${ROFI_THEME_FILE}`);
    makeFile(themePath);

    const result = withPlatform('linux', () => findRofiTheme('/usr/bin/subminer'));
    assert.equal(result, themePath);
  } finally {
    if (originalXdgDataHome !== undefined) {
      process.env.XDG_DATA_HOME = originalXdgDataHome;
    } else {
      delete process.env.XDG_DATA_HOME;
    }
    fs.rmSync(baseDir, { recursive: true, force: true });
  }
});

test('findRofiTheme resolves ~/.local/share/SubMiner/themes/subminer.rasi when XDG_DATA_HOME unset', () => {
  const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-home-'));
  const originalHomedir = os.homedir;
  const originalXdgDataHome = process.env.XDG_DATA_HOME;
  try {
    os.homedir = () => baseDir;
    delete process.env.XDG_DATA_HOME;
    const themePath = path.join(baseDir, `.local/share/SubMiner/themes/${ROFI_THEME_FILE}`);
    makeFile(themePath);

    const result = withPlatform('linux', () => findRofiTheme('/usr/bin/subminer'));
    assert.equal(result, themePath);
  } finally {
    os.homedir = originalHomedir;
    if (originalXdgDataHome !== undefined) {
      process.env.XDG_DATA_HOME = originalXdgDataHome;
    }
    fs.rmSync(baseDir, { recursive: true, force: true });
  }
});

test('findRofiThumbnailerDataRoot resolves the managed XDG data root', () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-xdg-'));
  const originalXdgDataHome = process.env.XDG_DATA_HOME;
  try {
    process.env.XDG_DATA_HOME = xdgDataHome;
    const dataRoot = path.join(xdgDataHome, 'SubMiner');
    makeFile(path.join(dataRoot, 'thumbnailers', ROFI_THUMBNAILER_FILE));

    const result = withPlatform('linux', () => findRofiThumbnailerDataRoot('/usr/bin/subminer'));
    assert.equal(result, dataRoot);
  } finally {
    if (originalXdgDataHome === undefined) {
      delete process.env.XDG_DATA_HOME;
    } else {
      process.env.XDG_DATA_HOME = originalXdgDataHome;
    }
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
  }
});

test('findRofiThumbnailerDataRoot is Linux-only', () => {
  assert.equal(
    withPlatform('darwin', () => findRofiThumbnailerDataRoot('/usr/bin/subminer')),
    null,
  );
});

test('prependXdgDataDir preserves existing roots and avoids duplicates', () => {
  const root = '/tmp/subminer-data';
  assert.equal(
    prependXdgDataDir(root, `/opt/share${path.delimiter}${root}${path.delimiter}/usr/share`),
    `${root}${path.delimiter}/opt/share${path.delimiter}/usr/share`,
  );
  assert.equal(
    prependXdgDataDir(root),
    `${root}${path.delimiter}/usr/local/share${path.delimiter}/usr/share`,
  );
});
