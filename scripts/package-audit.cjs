const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const asar = require('@electron/asar');
const { Arch } = require('builder-util');

const MIB = 1024 * 1024;
const currentReports = new Set();
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

// Do not follow framework symlinks or count ASAR unpacked entries twice.
function listFiles(root, prefix = '') {
  return fs.readdirSync(path.join(root, prefix), { withFileTypes: true }).flatMap((entry) => {
    const name = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isSymbolicLink()) return [];
    if (entry.isDirectory()) return listFiles(root, name);
    return [{ path: name, bytes: fs.statSync(path.join(root, name)).size }];
  });
}

function listAppFiles(archive) {
  return asar.listPackage(archive).flatMap((entry) => {
    const name = entry.replaceAll('\\', '/').replace(/^\//, '');
    const stat = asar.statFile(archive, name);
    return 'size' in stat ? [{ path: name, bytes: stat.size }] : [];
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
  const names = new Set(entries.map((entry) => entry.path));
  for (const name of REQUIRED_APP_FILES) assert(names.has(name), `Missing app file: ${name}`);
  for (const name of REQUIRED_RESOURCES) {
    assert(fs.statSync(path.join(resources, name)).size > 0, `Empty resource: ${name}`);
  }
  assert(listFiles(path.join(resources, 'yomitan-jlpt-vocab')).length > 0, 'Missing JLPT data');
  for (const { path: name } of entries) verifyAppPath(name, platform, arch);
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
    assert(!name.path.startsWith('minecard'), `Demo media shipped: ${name.path}`);
  }
  for (const ui of ['renderer', 'settings', 'syncui']) {
    const css = asar.extractFile(archive, `dist/${ui}/style.css`).toString();
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
  const appFiles = verifyContents(path.join(resources, 'app.asar'), resources, platform, arch);
  const files = listFiles(appRoot);
  const unpackedBytes = files.reduce((sum, entry) => sum + entry.bytes, 0);
  const report = {
    version: context.packager.appInfo.version,
    platform,
    arch,
    unpackedBytes,
    appDirectory: path.relative(context.outDir, appRoot),
    largestFiles: [...files].sort((a, b) => b.bytes - a.bytes).slice(0, 25),
    largestAppFiles: [...appFiles].sort((a, b) => b.bytes - a.bytes).slice(0, 25),
    nativeBinaries: files.filter((entry) => /\.(node|dll|dylib)$|\.so(?:\.|$)/.test(entry.path)),
    artifacts: [],
  };
  const output = path.join(context.outDir, `package-size-${key}.json`);
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, `${JSON.stringify(report, null, 2)}\n`);
  currentReports.add(output);
  console.log(
    `Package contents verified: ${key}, ${(unpackedBytes / MIB).toFixed(2)} MiB unpacked`,
  );
}

function artifactKind(name) {
  if (name.endsWith('-mac.zip')) return 'mac.zip';
  if (name.endsWith('-win.zip')) return 'win.zip';
  const extension = path.extname(name).slice(1);
  return ['AppImage', 'dmg', 'exe'].includes(extension) ? extension : undefined;
}

function compareSizes(report, previous) {
  assert.equal(previous.platform, report.platform);
  assert.equal(previous.arch, report.arch);
  assert(Number.isFinite(previous.unpackedBytes), 'Invalid previous size report');
  const previousArtifacts = Array.isArray(previous.artifacts) ? previous.artifacts : [];
  return {
    version: previous.version,
    unpackedDeltaBytes: report.unpackedBytes - previous.unpackedBytes,
    artifacts: report.artifacts.flatMap((artifact) => {
      const old = previousArtifacts.find(
        (entry) => entry && entry.kind === artifact.kind && Number.isFinite(entry.bytes),
      );
      return old ? [{ kind: artifact.kind, deltaBytes: artifact.bytes - old.bytes }] : [];
    }),
  };
}

// Runs after signing and installer creation, before release upload.
async function afterAllArtifactBuild(result) {
  const reports = [];
  for (const reportPath of currentReports) {
    const filename = path.basename(reportPath);
    const report = JSON.parse(fs.readFileSync(reportPath, 'utf8'));
    const key = `${report.platform}-${report.arch}`;
    const files = listFiles(path.join(result.outDir, report.appDirectory));
    report.unpackedBytes = files.reduce((sum, entry) => sum + entry.bytes, 0);
    report.largestFiles = [...files].sort((a, b) => b.bytes - a.bytes).slice(0, 25);
    report.artifacts = result.artifactPaths.flatMap((file) => {
      const kind = artifactKind(file);
      if (!kind) return [];
      const bytes = fs.statSync(file).size;
      return [{ name: path.basename(file), kind, bytes }];
    });
    const previousPath = path.join(result.outDir, '..', '.tmp', 'package-baseline', filename);
    if (fs.existsSync(previousPath)) {
      const previous = JSON.parse(fs.readFileSync(previousPath, 'utf8'));
      report.comparison = compareSizes(report, previous);
    }
    fs.writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`);
    const summary = [
      `### Package size: ${key}`,
      '',
      `Unpacked: ${(report.unpackedBytes / MIB).toFixed(2)} MiB`,
      ...report.artifacts.map((entry) => `${entry.name}: ${(entry.bytes / MIB).toFixed(2)} MiB`),
      report.comparison
        ? `Change from ${report.comparison.version}: ${(report.comparison.unpackedDeltaBytes / MIB).toFixed(2)} MiB unpacked`
        : 'No previous size report available.',
      '',
    ].join('\n');
    console.log(summary);
    if (process.env.GITHUB_STEP_SUMMARY)
      fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, summary);
    reports.push(reportPath);
  }
  assert(reports.length > 0, 'No package size reports generated by afterPack');
  return reports;
}

module.exports = {
  auditPackage,
  verifyContents,
  verifyAppPath,
  listFiles,
  listAppFiles,
  compareSizes,
  default: afterAllArtifactBuild,
};
