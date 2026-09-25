const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const asar = require('@electron/asar');
const { Arch } = require('builder-util');

const REQUIRED_APP_FILES = [
  'package.json',
  'LICENSE',
  'config.example.jsonc',
  'dist/main-entry.js',
  'dist/main.js',
  'dist/preload.js',
  'dist/preload-settings.js',
  'dist/preload-syncui.js',
  'dist/preload-stats.js',
  'dist/preload-jellyfin-setup.js',
  'dist/fonts/MPLUS1[wght].ttf',
  'stats/dist/index.html',
  'vendor/texthooker-ui/docs/index.html',
  ...['renderer', 'settings', 'syncui'].flatMap((ui) => [
    `dist/${ui}/index.html`,
    `dist/${ui}/style.css`,
    `dist/${ui}/${ui}.js`,
  ]),
];
const REQUIRED_RESOURCES = [
  'yomitan/manifest.json',
  'yomitan/data/fonts/kanji-stroke-orders.ttf',
  'yomitan/fonts/NotoSansJP-Regular.ttf',
  'yomitan/lib/resvg.wasm',
  'launcher/subminer',
  'plugin/subminer/main.lua',
  'plugin/subminer.conf',
  'assets/SubMiner.png',
  'assets/SubMiner-square.png',
  'assets/themes/subminer.rasi',
  'assets/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer',
  'CHANGELOG.md',
];

// Skip symlinks when checking resource contents.
function listFiles(root, prefix = '') {
  return fs.readdirSync(path.join(root, prefix), { withFileTypes: true }).flatMap((entry) => {
    const name = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isSymbolicLink()) return [];
    if (entry.isDirectory()) return listFiles(root, name);
    return [name];
  });
}

// asar resolves lookups with the platform separator, so stat with the listed
// native path and only normalize the reported name.
function listAppFiles(archive) {
  return asar.listPackage(archive).flatMap((entry) => {
    const native = entry.replace(/^[\\/]/, '');
    const stat = asar.statFile(archive, native);
    const name = native.replaceAll('\\', '/');
    return 'size' in stat ? [name] : [];
  });
}

function verifyAppPath(name, platform, arch) {
  const allowedRoots = new Set([
    'dist',
    'node_modules',
    'stats',
    'vendor',
    'package.json',
    'LICENSE',
    'config.example.jsonc',
  ]);
  assert(allowedRoots.has(name.split('/')[0]), `Unexpected app file: ${name}`);
  assert(!name.endsWith('.map'), `Packaged source map: ${name}`);
  assert(!/\.(?:[cm]?ts|tsx)$/.test(name), `Packaged TypeScript: ${name}`);
  assert(!/\.(?:test|spec)\./.test(name), `Packaged test: ${name}`);
  assert(
    !/(?:^|\/)(?:tests?|__tests__|fixtures?|__fixtures__)\//.test(name),
    `Packaged test or fixture directory: ${name}`,
  );
  assert(!/^dist\/.*\.test\./.test(name), `Packaged test: ${name}`);
  assert(!/^dist\/(launcher|scripts)\//.test(name), `Duplicate helper: ${name}`);
  assert(!/^dist\/(renderer|settings|syncui)\/fonts\//.test(name), `Duplicate font: ${name}`);
  assert(!name.startsWith('stats/') || name.startsWith('stats/dist/'), `Stats source: ${name}`);
  assert(
    !name.startsWith('vendor/') || name.startsWith('vendor/texthooker-ui/docs/'),
    `Vendor source: ${name}`,
  );
  if (name.startsWith('node_modules/koffi/')) {
    assert.equal(platform, 'win32', `Koffi shipped on ${platform}`);
    assert(!/^node_modules\/koffi\/(src|vendor|doc)\//.test(name), `Koffi build files: ${name}`);
    if (name.endsWith('.node')) {
      assert.equal(name, `node_modules/koffi/build/koffi/win32_${arch}/koffi.node`);
    }
  }
}

function verifyContents(archive, resources, platform, arch) {
  const entries = listAppFiles(archive);
  const names = new Set(entries);
  for (const name of REQUIRED_APP_FILES) assert(names.has(name), `Missing app file: ${name}`);
  for (const name of REQUIRED_RESOURCES) {
    assert(fs.statSync(path.join(resources, name)).size > 0, `Empty resource: ${name}`);
  }
  assert(listFiles(path.join(resources, 'yomitan-jlpt-vocab')).length > 0, 'Missing JLPT data');
  for (const name of entries) verifyAppPath(name, platform, arch);
  const libsqlPlatform = {
    linux: `linux-${arch}-gnu`,
    darwin: `darwin-${arch}`,
    win32: `win32-${arch}-msvc`,
  }[platform];
  const libsqlBinary = `node_modules/@libsql/${libsqlPlatform}/index.node`;
  assert(names.has(libsqlBinary), `Missing SQLite native binary: ${libsqlBinary}`);
  for (const name of names) {
    if (name.startsWith('node_modules/@libsql/') && name.endsWith('.node')) {
      assert.equal(name, libsqlBinary, `Foreign SQLite binary: ${name}`);
    }
  }
  if (platform === 'win32') {
    for (const name of [
      'index.js',
      'package.json',
      'LICENSE.txt',
      `build/koffi/win32_${arch}/koffi.node`,
    ]) {
      assert(names.has(`node_modules/koffi/${name}`), `Missing Windows FFI file: ${name}`);
    }
  }
  for (const name of listFiles(path.join(resources, 'assets'))) {
    assert(!name.startsWith('minecard'), `Demo media shipped: ${name}`);
  }
  for (const ui of ['renderer', 'settings', 'syncui']) {
    const css = asar.extractFile(archive, path.join('dist', ui, 'style.css')).toString();
    assert(css.includes('../fonts/MPLUS1[wght].ttf'), `Shared font missing from ${ui} CSS`);
  }
  return entries;
}

async function auditPackage(context) {
  const platform = context.electronPlatformName;
  const arch = Arch[context.arch];
  const key = `${platform}-${arch}`;
  const appRoot =
    platform === 'darwin'
      ? path.join(context.appOutDir, `${context.packager.appInfo.productFilename}.app`)
      : context.appOutDir;
  const resources = path.join(appRoot, platform === 'darwin' ? 'Contents/Resources' : 'resources');
  verifyContents(path.join(resources, 'app.asar'), resources, platform, arch);
  console.log(`Package contents verified: ${key}`);
}

module.exports = {
  auditPackage,
  verifyContents,
  verifyAppPath,
  listFiles,
  listAppFiles,
};
