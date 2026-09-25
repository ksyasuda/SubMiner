import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  ensureLinuxRuntimePluginAssets,
  resolveManagedLinuxRuntimePluginPaths,
} from './linux-runtime-plugin-assets';

const THUMBNAILER_RELATIVE_PATH = path.join(
  'thumbnailers',
  'subminer-ffmpegthumbnailer.thumbnailer',
);

function writeThumbnailer(rootDir: string, content = '[Thumbnailer Entry]\n'): string {
  const thumbnailerPath = path.join(rootDir, THUMBNAILER_RELATIVE_PATH);
  fs.mkdirSync(path.dirname(thumbnailerPath), { recursive: true });
  fs.writeFileSync(thumbnailerPath, content);
  return thumbnailerPath;
}

async function withTempDir<T>(fn: (dir: string) => Promise<T> | T): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-linux-plugin-assets-test-'));
  try {
    return await fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

async function withProcessResourcesPath<T>(
  resourcesPath: string,
  fn: () => Promise<T> | T,
): Promise<T> {
  const processWithResources = process as NodeJS.Process & { resourcesPath?: string };
  const previousResourcesPath = processWithResources.resourcesPath;
  processWithResources.resourcesPath = resourcesPath;
  try {
    return await fn();
  } finally {
    if (previousResourcesPath === undefined) {
      Reflect.deleteProperty(processWithResources, 'resourcesPath');
    } else {
      processWithResources.resourcesPath = previousResourcesPath;
    }
  }
}

test('resolveManagedLinuxRuntimePluginPaths resolves XDG data target paths', () => {
  const resolved = resolveManagedLinuxRuntimePluginPaths({
    homeDir: '/home/tester',
    xdgDataHome: '/tmp/xdg-data',
  });

  assert.deepEqual(resolved, {
    dataDir: '/tmp/xdg-data/SubMiner',
    rootDir: '/tmp/xdg-data/SubMiner/plugin',
    pluginDir: '/tmp/xdg-data/SubMiner/plugin/subminer',
    pluginEntrypointPath: '/tmp/xdg-data/SubMiner/plugin/subminer/main.lua',
    pluginConfigPath: '/tmp/xdg-data/SubMiner/plugin/subminer.conf',
    themePath: '/tmp/xdg-data/SubMiner/themes/subminer.rasi',
    thumbnailerPath: '/tmp/xdg-data/SubMiner/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer',
  });
});

test('resolveManagedLinuxRuntimePluginPaths treats blank XDG data homes as unset', () => {
  const previousXdgDataHome = process.env.XDG_DATA_HOME;
  try {
    delete process.env.XDG_DATA_HOME;
    const emptyOption = resolveManagedLinuxRuntimePluginPaths({
      homeDir: '/home/tester',
      xdgDataHome: '',
    });
    assert.equal(emptyOption.dataDir, path.join('/home/tester', '.local', 'share', 'SubMiner'));

    process.env.XDG_DATA_HOME = '   ';
    const blankEnv = resolveManagedLinuxRuntimePluginPaths({
      homeDir: '/home/tester',
    });
    assert.equal(blankEnv.dataDir, path.join('/home/tester', '.local', 'share', 'SubMiner'));
  } finally {
    if (previousXdgDataHome === undefined) {
      delete process.env.XDG_DATA_HOME;
    } else {
      process.env.XDG_DATA_HOME = previousXdgDataHome;
    }
  }
});

test('ensureLinuxRuntimePluginAssets installs managed plugin dir, config, and rofi theme when missing', async () => {
  await withTempDir(async (tempDir) => {
    const sourceRoot = path.join(tempDir, 'source', 'plugin');
    const themeSourcePath = path.join(tempDir, 'source', 'assets', 'themes', 'subminer.rasi');
    const thumbnailerSourcePath = writeThumbnailer(path.join(tempDir, 'source', 'assets'));
    const targetRoot = path.join(tempDir, 'xdg-data', 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(sourceRoot, 'subminer'), { recursive: true });
    fs.mkdirSync(path.dirname(themeSourcePath), { recursive: true });
    fs.writeFileSync(path.join(sourceRoot, 'subminer', 'main.lua'), '-- plugin\n');
    fs.writeFileSync(path.join(sourceRoot, 'subminer.conf'), 'configured=true\n');
    fs.writeFileSync(themeSourcePath, '/* theme */\n');

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome: path.join(tempDir, 'xdg-data'),
      resolveBundledAssets: () => ({
        pluginDirSource: path.join(sourceRoot, 'subminer'),
        pluginConfigSource: path.join(sourceRoot, 'subminer.conf'),
        themeSourcePath,
        thumbnailerSourcePath,
      }),
    });

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer', 'main.lua'), 'utf8'),
      '-- plugin\n',
    );
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer.conf'), 'utf8'),
      'configured=true\n',
    );
    assert.equal(
      fs.readFileSync(
        path.join(tempDir, 'xdg-data', 'SubMiner', 'themes', 'subminer.rasi'),
        'utf8',
      ),
      '/* theme */\n',
    );
    assert.equal(
      fs.readFileSync(
        path.join(
          tempDir,
          'xdg-data',
          'SubMiner',
          'thumbnailers',
          'subminer-ffmpegthumbnailer.thumbnailer',
        ),
        'utf8',
      ),
      '[Thumbnailer Entry]\n',
    );
  });
});

test('ensureLinuxRuntimePluginAssets installs managed theme when plugin assets already exist', async () => {
  await withTempDir(async (tempDir) => {
    const sourceRoot = path.join(tempDir, 'source', 'plugin');
    const themeSourcePath = path.join(tempDir, 'source', 'assets', 'themes', 'subminer.rasi');
    const thumbnailerSourcePath = writeThumbnailer(path.join(tempDir, 'source', 'assets'));
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(sourceRoot, 'subminer'), { recursive: true });
    fs.mkdirSync(path.dirname(themeSourcePath), { recursive: true });
    fs.mkdirSync(path.join(targetRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(sourceRoot, 'subminer', 'main.lua'), '-- new plugin\n');
    fs.writeFileSync(path.join(sourceRoot, 'subminer.conf'), 'new=true\n');
    fs.writeFileSync(themeSourcePath, '/* theme */\n');
    fs.writeFileSync(path.join(targetRoot, 'subminer', 'main.lua'), '-- existing plugin\n');
    fs.writeFileSync(path.join(targetRoot, 'subminer.conf'), 'configured=true\n');

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome,
      resolveBundledAssets: () => ({
        pluginDirSource: path.join(sourceRoot, 'subminer'),
        pluginConfigSource: path.join(sourceRoot, 'subminer.conf'),
        themeSourcePath,
        thumbnailerSourcePath,
      }),
    });

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer', 'main.lua'), 'utf8'),
      '-- existing plugin\n',
    );
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer.conf'), 'utf8'),
      'configured=true\n',
    );
    assert.equal(
      fs.readFileSync(path.join(xdgDataHome, 'SubMiner', 'themes', 'subminer.rasi'), 'utf8'),
      '/* theme */\n',
    );
  });
});

test('ensureLinuxRuntimePluginAssets installs managed theme without resolving plugin sources when plugin assets already exist', async () => {
  await withTempDir(async (tempDir) => {
    const themeSourcePath = path.join(tempDir, 'source', 'assets', 'themes', 'subminer.rasi');
    const thumbnailerSourcePath = writeThumbnailer(path.join(tempDir, 'source', 'assets'));
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.dirname(themeSourcePath), { recursive: true });
    fs.mkdirSync(path.join(targetRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(themeSourcePath, '/* theme */\n');
    fs.writeFileSync(path.join(targetRoot, 'subminer', 'main.lua'), '-- existing plugin\n');
    fs.writeFileSync(path.join(targetRoot, 'subminer.conf'), 'configured=true\n');

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome,
      resolveBundledAssets: () => ({
        themeSourcePath,
        thumbnailerSourcePath,
      }),
    });

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
    assert.equal(
      fs.readFileSync(path.join(xdgDataHome, 'SubMiner', 'themes', 'subminer.rasi'), 'utf8'),
      '/* theme */\n',
    );
  });
});

test('ensureLinuxRuntimePluginAssets default resolver can recover a missing theme when plugin sources are unavailable', async () => {
  await withTempDir(async (tempDir) => {
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(targetRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(targetRoot, 'subminer', 'main.lua'), '-- existing plugin\n');
    fs.writeFileSync(path.join(targetRoot, 'subminer.conf'), 'configured=true\n');

    const result = await withProcessResourcesPath(process.cwd(), () =>
      ensureLinuxRuntimePluginAssets({
        platform: 'linux',
        homeDir: path.join(tempDir, 'home'),
        xdgDataHome,
        existsSync: (candidate) => {
          if (
            !candidate.startsWith(tempDir) &&
            (candidate.endsWith(path.join('plugin', 'subminer')) ||
              candidate.endsWith(path.join('plugin', 'subminer.conf')) ||
              candidate.endsWith(path.join('plugin', 'subminer', 'main.lua')))
          ) {
            return false;
          }
          return fs.existsSync(candidate);
        },
      }),
    );

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
    assert.equal(
      fs.readFileSync(path.join(xdgDataHome, 'SubMiner', 'themes', 'subminer.rasi'), 'utf8'),
      fs.readFileSync(path.join(process.cwd(), 'assets', 'themes', 'subminer.rasi'), 'utf8'),
    );
  });
});

test('ensureLinuxRuntimePluginAssets installs managed plugin assets without resolving theme source when theme already exists', async () => {
  await withTempDir(async (tempDir) => {
    const sourceRoot = path.join(tempDir, 'source', 'plugin');
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(sourceRoot, 'subminer'), { recursive: true });
    fs.mkdirSync(path.join(xdgDataHome, 'SubMiner', 'themes'), { recursive: true });
    fs.writeFileSync(path.join(sourceRoot, 'subminer', 'main.lua'), '-- plugin\n');
    fs.writeFileSync(path.join(sourceRoot, 'subminer.conf'), 'configured=true\n');
    fs.writeFileSync(
      path.join(xdgDataHome, 'SubMiner', 'themes', 'subminer.rasi'),
      '/* existing theme */\n',
    );
    const thumbnailerSourcePath = writeThumbnailer(path.join(tempDir, 'source', 'assets'));

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome,
      resolveBundledAssets: () => ({
        pluginDirSource: path.join(sourceRoot, 'subminer'),
        pluginConfigSource: path.join(sourceRoot, 'subminer.conf'),
        thumbnailerSourcePath,
      }),
    });

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer', 'main.lua'), 'utf8'),
      '-- plugin\n',
    );
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer.conf'), 'utf8'),
      'configured=true\n',
    );
  });
});

test('ensureLinuxRuntimePluginAssets default resolver can recover missing plugin assets when theme source is unavailable', async () => {
  await withTempDir(async (tempDir) => {
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(xdgDataHome, 'SubMiner', 'themes'), { recursive: true });
    fs.writeFileSync(
      path.join(xdgDataHome, 'SubMiner', 'themes', 'subminer.rasi'),
      '/* existing theme */\n',
    );

    const result = await withProcessResourcesPath(process.cwd(), () =>
      ensureLinuxRuntimePluginAssets({
        platform: 'linux',
        homeDir: path.join(tempDir, 'home'),
        xdgDataHome,
        existsSync: (candidate) => {
          if (
            !candidate.startsWith(tempDir) &&
            candidate.endsWith(path.join('assets', 'themes', 'subminer.rasi'))
          ) {
            return false;
          }
          return fs.existsSync(candidate);
        },
      }),
    );

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer', 'main.lua'), 'utf8'),
      fs.readFileSync(path.join(process.cwd(), 'plugin', 'subminer', 'main.lua'), 'utf8'),
    );
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer.conf'), 'utf8'),
      fs.readFileSync(path.join(process.cwd(), 'plugin', 'subminer.conf'), 'utf8'),
    );
  });
});

test('ensureLinuxRuntimePluginAssets returns already-present when managed assets already exist', async () => {
  await withTempDir(async (tempDir) => {
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(targetRoot, 'subminer'), { recursive: true });
    fs.mkdirSync(path.join(xdgDataHome, 'SubMiner', 'themes'), { recursive: true });
    fs.writeFileSync(path.join(targetRoot, 'subminer', 'main.lua'), '-- existing\n');
    fs.writeFileSync(path.join(targetRoot, 'subminer.conf'), 'configured=true\n');
    fs.writeFileSync(
      path.join(xdgDataHome, 'SubMiner', 'themes', 'subminer.rasi'),
      '/* theme */\n',
    );
    writeThumbnailer(path.join(xdgDataHome, 'SubMiner'));

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome,
      resolveBundledAssets: () => {
        throw new Error('should not resolve bundled assets when already installed');
      },
    });

    assert.deepEqual(result, {
      ok: true,
      status: 'already-present',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
  });
});

test('ensureLinuxRuntimePluginAssets fails when bundled assets cannot be resolved', async () => {
  await withTempDir(async (tempDir) => {
    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome: path.join(tempDir, 'xdg-data'),
      resolveBundledAssets: () => ({}),
    });

    assert.equal(result.ok, false);
    assert.equal(result.status, 'failed');
    assert.match(result.error ?? '', /bundled.*plugin assets/i);
  });
});

test('ensureLinuxRuntimePluginAssets leaves no final target tree on failed install', async () => {
  await withTempDir(async (tempDir) => {
    const sourceRoot = path.join(tempDir, 'source', 'plugin');
    const themeSourcePath = path.join(tempDir, 'source', 'assets', 'themes', 'subminer.rasi');
    const thumbnailerSourcePath = writeThumbnailer(path.join(tempDir, 'source', 'assets'));
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(sourceRoot, 'subminer'), { recursive: true });
    fs.mkdirSync(path.dirname(themeSourcePath), { recursive: true });
    fs.writeFileSync(path.join(sourceRoot, 'subminer', 'main.lua'), '-- plugin\n');
    fs.writeFileSync(path.join(sourceRoot, 'subminer.conf'), 'configured=true\n');
    fs.writeFileSync(themeSourcePath, '/* theme */\n');

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome,
      resolveBundledAssets: () => ({
        pluginDirSource: path.join(sourceRoot, 'subminer'),
        pluginConfigSource: path.join(sourceRoot, 'subminer.conf'),
        themeSourcePath,
        thumbnailerSourcePath,
      }),
      copyFile: async () => {
        throw new Error('copy failed');
      },
    });

    assert.equal(result.ok, false);
    assert.equal(result.status, 'failed');
    assert.equal(fs.existsSync(path.join(targetRoot, 'subminer', 'main.lua')), false);
    assert.equal(fs.existsSync(path.join(targetRoot, 'subminer.conf')), false);
  });
});
