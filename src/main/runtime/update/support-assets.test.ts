import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  buildProtectedSupportAssetsCommand,
  detectSupportAssetDataDirs,
  updateSupportAssetsFromRelease,
  type SupportAssetsUpdateResult,
} from './support-assets';

type SupportAssetsResultWithComponent = SupportAssetsUpdateResult & {
  component?: 'theme' | 'plugin';
};

function sha256(data: Buffer): string {
  return createHash('sha256').update(data).digest('hex');
}

function makeSupportAssetsArchive(options?: {
  themeContent?: string;
  pluginVersion?: string | null;
  pluginMainContent?: string;
  extraPluginFiles?: Array<{ relativePath: string; content: string }>;
}): { archive: Buffer; tempDir: string } {
  const themeContent = options?.themeContent ?? 'new theme\n';
  const pluginVersion = options && 'pluginVersion' in options ? options.pluginVersion : '0.12.0';
  const pluginMainContent = options?.pluginMainContent ?? 'new plugin\n';
  const extraPluginFiles = options?.extraPluginFiles ?? [];
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-support-assets-test-'));
  fs.mkdirSync(path.join(tempDir, 'assets/themes'), { recursive: true });
  fs.mkdirSync(path.join(tempDir, 'plugin/subminer'), { recursive: true });
  fs.writeFileSync(path.join(tempDir, 'assets/themes/subminer.rasi'), themeContent);
  fs.writeFileSync(path.join(tempDir, 'plugin/subminer/main.lua'), pluginMainContent);
  if (pluginVersion !== null) {
    fs.writeFileSync(
      path.join(tempDir, 'plugin/subminer/version.lua'),
      `return {\n\tversion = "${pluginVersion}",\n}\n`,
    );
  }
  for (const extraFile of extraPluginFiles) {
    const targetPath = path.join(tempDir, 'plugin/subminer', extraFile.relativePath);
    fs.mkdirSync(path.dirname(targetPath), { recursive: true });
    fs.writeFileSync(targetPath, extraFile.content);
  }
  execFileSync('tar', ['-czf', 'subminer-assets.tar.gz', 'assets', 'plugin'], { cwd: tempDir });
  return {
    archive: fs.readFileSync(path.join(tempDir, 'subminer-assets.tar.gz')),
    tempDir,
  };
}

async function runLinuxSupportAssetUpdate(options: {
  archive: Buffer;
  xdgDataHome?: string;
  platform?: NodeJS.Platform;
}): Promise<SupportAssetsResultWithComponent[]> {
  return (await updateSupportAssetsFromRelease({
    release: {
      tag_name: 'v0.15.0',
      assets: [
        {
          name: 'subminer-assets.tar.gz',
          browser_download_url: 'https://example.test/subminer-assets.tar.gz',
        },
      ],
    },
    sha256Sums: new Map([['subminer-assets.tar.gz', sha256(options.archive)]]),
    downloadAsset: async () => options.archive,
    platform: options.platform ?? 'linux',
    xdgDataHome: options.xdgDataHome,
  })) as SupportAssetsResultWithComponent[];
}

test('detectSupportAssetDataDirs only returns Linux support-asset locations', () => {
  assert.deepEqual(
    detectSupportAssetDataDirs({
      platform: 'darwin',
      homeDir: '/Users/kyle',
    }),
    [],
  );
  assert.deepEqual(
    detectSupportAssetDataDirs({
      platform: 'linux',
      homeDir: '/home/kyle',
      xdgDataHome: '/tmp/xdg-data',
    }),
    ['/tmp/xdg-data/SubMiner', '/usr/local/share/SubMiner', '/usr/share/SubMiner'],
  );
});

test('buildProtectedSupportAssetsCommand installs both theme and plugin assets', () => {
  const command = buildProtectedSupportAssetsCommand(
    "https://example.test/subminer assets.tar.gz?sig='abc'",
    "/usr/local/share/SubMiner's data",
  );

  assert.match(command, /tmp=\$\(mktemp -d\)/);
  assert.match(command, /trap 'rm -rf "\$tmp"' EXIT/);
  assert.match(
    command,
    /curl -fSL 'https:\/\/example\.test\/subminer assets\.tar\.gz\?sig='\\''abc'\\''' -o "\$tmp\/subminer-assets\.tar\.gz"/,
  );
  assert.match(command, /sudo mkdir -p '\/usr\/local\/share\/SubMiner'\\''s data'\/themes/);
  assert.match(
    command,
    /sudo cp "\$tmp\/assets\/themes\/subminer\.rasi" '\/usr\/local\/share\/SubMiner'\\''s data'\/themes\/subminer\.rasi/,
  );
  assert.match(command, /sudo mkdir -p '\/usr\/local\/share\/SubMiner'\\''s data'\/plugin/);
  assert.match(command, /sudo rm -rf .*plugin\/subminer\.next/);
  assert.match(command, /sudo cp -R "\$tmp\/plugin\/subminer" .*plugin\/subminer\.next/);
  assert.match(command, /sudo mv .*plugin\/subminer\.next.*plugin\/subminer/);
});

test('updateSupportAssetsFromRelease skips on non-Linux platforms', async () => {
  const { archive, tempDir } = makeSupportAssetsArchive();
  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      platform: 'darwin',
    });

    assert.deepEqual(results, [
      {
        status: 'skipped',
        message: 'Support assets are only installed on Linux.',
      },
    ]);
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease skips when no managed support-asset roots exist', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const { archive, tempDir } = makeSupportAssetsArchive();

  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      xdgDataHome,
    });

    assert.deepEqual(results, [
      {
        status: 'skipped',
        message: 'No managed SubMiner support-asset install detected.',
      },
    ]);
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease skips existing data roots without managed asset markers', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.posix.join(xdgDataHome, 'SubMiner');
  fs.mkdirSync(dataDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'preferences.json'), '{}\n');
  const { archive, tempDir } = makeSupportAssetsArchive();

  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      xdgDataHome,
    });

    assert.deepEqual(results, [
      {
        status: 'skipped',
        message: 'No managed SubMiner support-asset install detected.',
      },
    ]);
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease installs missing plugin into a root with a managed theme marker', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.posix.join(xdgDataHome, 'SubMiner');
  fs.mkdirSync(path.join(dataDir, 'themes'), { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'old theme\n');
  const { archive, tempDir } = makeSupportAssetsArchive({
    extraPluginFiles: [{ relativePath: 'nested.lua', content: 'nested file\n' }],
  });

  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      xdgDataHome,
    });

    assert.deepEqual(results, [
      {
        status: 'updated',
        component: 'theme',
        path: dataDir,
        message: 'Updated theme.',
      },
      {
        status: 'updated',
        component: 'plugin',
        path: dataDir,
        message: 'Installed plugin.',
      },
    ]);
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'utf8'),
      'new theme\n',
    );
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'plugin/subminer/main.lua'), 'utf8'),
      'new plugin\n',
    );
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'plugin/subminer/version.lua'), 'utf8'),
      'return {\n\tversion = "0.12.0",\n}\n',
    );
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'plugin/subminer/nested.lua'), 'utf8'),
      'nested file\n',
    );
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease preserves existing plugin when staged replacement copy fails', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.posix.join(xdgDataHome, 'SubMiner');
  const pluginDir = path.join(dataDir, 'plugin/subminer');
  fs.mkdirSync(path.join(dataDir, 'themes'), { recursive: true });
  fs.mkdirSync(pluginDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'same theme\n');
  fs.writeFileSync(path.join(pluginDir, 'main.lua'), 'old plugin\n');
  fs.writeFileSync(path.join(pluginDir, 'version.lua'), 'return {\n\tversion = "0.11.0",\n}\n');
  const { archive, tempDir } = makeSupportAssetsArchive({
    themeContent: 'same theme\n',
    pluginVersion: '0.12.0',
    extraPluginFiles: [{ relativePath: 'blocked.lua', content: 'blocked\n' }],
  });

  try {
    const originalCp = fs.promises.cp;
    fs.promises.cp = async (...args: Parameters<typeof fs.promises.cp>) => {
      const targetPath = String(args[1]);
      if (
        targetPath.endsWith(`${path.sep}subminer`) ||
        targetPath.endsWith(`${path.sep}subminer.next`)
      ) {
        throw new Error('copy failed');
      }
      return originalCp(...args);
    };
    try {
      await assert.rejects(
        () =>
          runLinuxSupportAssetUpdate({
            archive,
            xdgDataHome,
          }),
        /copy failed/,
      );
    } finally {
      fs.promises.cp = originalCp;
    }

    assert.equal(fs.readFileSync(path.join(pluginDir, 'main.lua'), 'utf8'), 'old plugin\n');
    assert.equal(
      fs.readFileSync(path.join(pluginDir, 'version.lua'), 'utf8'),
      'return {\n\tversion = "0.11.0",\n}\n',
    );
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease skips identical theme and up-to-date plugin', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.posix.join(xdgDataHome, 'SubMiner');
  fs.mkdirSync(path.join(dataDir, 'themes'), { recursive: true });
  fs.mkdirSync(path.join(dataDir, 'plugin/subminer'), { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'same theme\n');
  fs.writeFileSync(path.join(dataDir, 'plugin/subminer/main.lua'), 'same plugin\n');
  fs.writeFileSync(
    path.join(dataDir, 'plugin/subminer/version.lua'),
    'return {\n\tversion = "0.12.0",\n}\n',
  );
  const { archive, tempDir } = makeSupportAssetsArchive({
    themeContent: 'same theme\n',
    pluginVersion: '0.12.0',
    pluginMainContent: 'release plugin differs but version matches\n',
  });

  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      xdgDataHome,
    });

    assert.deepEqual(results, [
      {
        status: 'skipped',
        component: 'theme',
        path: dataDir,
        message: 'Theme already up to date.',
      },
      {
        status: 'skipped',
        component: 'plugin',
        path: dataDir,
        message: 'Plugin already up to date.',
      },
    ]);
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'utf8'),
      'same theme\n',
    );
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'plugin/subminer/main.lua'), 'utf8'),
      'same plugin\n',
    );
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease updates changed theme and outdated plugin while removing stale plugin files', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.posix.join(xdgDataHome, 'SubMiner');
  fs.mkdirSync(path.join(dataDir, 'themes'), { recursive: true });
  fs.mkdirSync(path.join(dataDir, 'plugin/subminer'), { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'old theme\n');
  fs.writeFileSync(path.join(dataDir, 'plugin/subminer/main.lua'), 'old plugin\n');
  fs.writeFileSync(
    path.join(dataDir, 'plugin/subminer/version.lua'),
    'return {\n\tversion = "0.11.0",\n}\n',
  );
  fs.writeFileSync(path.join(dataDir, 'plugin/subminer/stale.lua'), 'stale\n');
  const { archive, tempDir } = makeSupportAssetsArchive({
    themeContent: 'new theme\n',
    pluginVersion: '0.12.0',
    pluginMainContent: 'new plugin main\n',
    extraPluginFiles: [{ relativePath: 'fresh.lua', content: 'fresh\n' }],
  });

  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      xdgDataHome,
    });

    assert.deepEqual(results, [
      {
        status: 'updated',
        component: 'theme',
        path: dataDir,
        message: 'Updated theme.',
      },
      {
        status: 'updated',
        component: 'plugin',
        path: dataDir,
        message: 'Updated plugin.',
      },
    ]);
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'utf8'),
      'new theme\n',
    );
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'plugin/subminer/main.lua'), 'utf8'),
      'new plugin main\n',
    );
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'plugin/subminer/fresh.lua'), 'utf8'),
      'fresh\n',
    );
    assert.equal(fs.existsSync(path.join(dataDir, 'plugin/subminer/stale.lua')), false);
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease returns protected commands for managed roots that are not writable', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.posix.join(xdgDataHome, 'SubMiner');
  fs.mkdirSync(path.join(dataDir, 'themes'), { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'managed theme\n');
  const originalMode = fs.statSync(dataDir).mode & 0o777;
  fs.chmodSync(dataDir, 0o555);
  const { archive, tempDir } = makeSupportAssetsArchive();

  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      xdgDataHome,
    });

    assert.deepEqual(
      results.map((result) => ({
        status: result.status,
        component: result.component,
        path: result.path,
        command: typeof result.command === 'string',
      })),
      [
        {
          status: 'protected',
          component: 'theme',
          path: dataDir,
          command: true,
        },
        {
          status: 'protected',
          component: 'plugin',
          path: dataDir,
          command: true,
        },
      ],
    );
    assert.match(results[0]?.command ?? '', /themes\/subminer\.rasi/);
    assert.match(results[0]?.command ?? '', /plugin\/subminer/);
  } finally {
    fs.chmodSync(dataDir, originalMode);
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('updateSupportAssetsFromRelease returns missing-asset when release plugin version metadata is absent', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.posix.join(xdgDataHome, 'SubMiner');
  fs.mkdirSync(path.join(dataDir, 'themes'), { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'managed theme\n');
  const { archive, tempDir } = makeSupportAssetsArchive({
    pluginVersion: null,
  });

  try {
    const results = await runLinuxSupportAssetUpdate({
      archive,
      xdgDataHome,
    });

    assert.deepEqual(results, [
      {
        status: 'missing-asset',
        message: 'Support asset archive has no readable plugin version metadata.',
      },
    ]);
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});
