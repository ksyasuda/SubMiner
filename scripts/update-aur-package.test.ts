import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

function createWorkspace(name: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), `${name}-`));
}

test('update-aur-package updates PKGBUILD and .SRCINFO without makepkg', () => {
  const workspace = createWorkspace('subminer-aur-package');
  const pkgDir = path.join(workspace, 'aur-subminer-bin');
  const appImagePath = path.join(workspace, 'SubMiner-0.6.3.AppImage');
  const wrapperPath = path.join(workspace, 'subminer');
  const assetsPath = path.join(workspace, 'subminer-assets.tar.gz');

  fs.mkdirSync(pkgDir, { recursive: true });
  fs.copyFileSync('packaging/aur/subminer-bin/PKGBUILD', path.join(pkgDir, 'PKGBUILD'));
  fs.copyFileSync('packaging/aur/subminer-bin/.SRCINFO', path.join(pkgDir, '.SRCINFO'));
  fs.writeFileSync(appImagePath, 'appimage');
  fs.writeFileSync(wrapperPath, 'wrapper');
  fs.writeFileSync(assetsPath, 'assets');

  try {
    execFileSync(
      'bash',
      [
        'scripts/update-aur-package.sh',
        '--pkg-dir',
        pkgDir,
        '--version',
        'v0.6.3',
        '--appimage',
        appImagePath,
        '--wrapper',
        wrapperPath,
        '--assets',
        assetsPath,
      ],
      {
        cwd: process.cwd(),
        encoding: 'utf8',
      },
    );

    const pkgbuild = fs.readFileSync(path.join(pkgDir, 'PKGBUILD'), 'utf8');
    const srcinfo = fs.readFileSync(path.join(pkgDir, '.SRCINFO'), 'utf8');
    const expectedSums = [appImagePath, wrapperPath, assetsPath].map((filePath) =>
      execFileSync('sha256sum', [filePath], { encoding: 'utf8' }).split(/\s+/)[0],
    );

    assert.match(pkgbuild, /^pkgver=0\.6\.3$/m);
    assert.match(
      pkgbuild,
      /^\s*"subminer-\$\{pkgver\}::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v\$\{pkgver\}\/subminer"$/m,
    );
    assert.match(
      pkgbuild,
      /^\s*"subminer-assets-\$\{pkgver\}\.tar\.gz::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v\$\{pkgver\}\/subminer-assets\.tar\.gz"$/m,
    );
    assert.match(
      pkgbuild,
      /^\s*install -Dm755 "\$\{srcdir\}\/subminer-\$\{pkgver\}" "\$\{pkgdir\}\/usr\/bin\/subminer"$/m,
    );
    assert.match(srcinfo, /^\tpkgver = 0\.6\.3$/m);
    assert.match(srcinfo, /^\tprovides = subminer=0\.6\.3$/m);
    assert.match(
      srcinfo,
      /^\tsource = SubMiner-0\.6\.3\.AppImage::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v0\.6\.3\/SubMiner-0\.6\.3\.AppImage$/m,
    );
    assert.match(
      srcinfo,
      /^\tsource = subminer-0\.6\.3::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v0\.6\.3\/subminer$/m,
    );
    assert.match(
      srcinfo,
      /^\tsource = subminer-assets-0\.6\.3\.tar\.gz::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v0\.6\.3\/subminer-assets\.tar\.gz$/m,
    );
    assert.match(srcinfo, new RegExp(`^\\tsha256sums = ${expectedSums[0]}$`, 'm'));
    assert.match(srcinfo, new RegExp(`^\\tsha256sums = ${expectedSums[1]}$`, 'm'));
    assert.match(srcinfo, new RegExp(`^\\tsha256sums = ${expectedSums[2]}$`, 'm'));
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});
