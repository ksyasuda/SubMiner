import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import packageJson from '../package.json';
import { posixLauncherBootstrapContent } from '../src/main/runtime/posix-launcher-bootstrap';
import { windowsLauncherBootstrapContent } from '../src/main/runtime/windows-launcher-bootstrap';

const outputDirectory = path.join(process.cwd(), 'dist', 'launcher');

function bundle(options: {
  entrypoint: string;
  outfile: string;
  target: 'bun' | 'node';
  format?: 'cjs';
  banner?: string;
}): void {
  const args = [
    'build',
    options.entrypoint,
    `--outfile=${options.outfile}`,
    `--target=${options.target}`,
    '--packages=bundle',
  ];
  if (options.format) args.push(`--format=${options.format}`);
  if (options.banner) args.push(`--banner=${options.banner}`);
  execFileSync(process.execPath, args, { stdio: 'inherit' });
}

fs.mkdirSync(outputDirectory, { recursive: true });

bundle({
  entrypoint: path.join(process.cwd(), 'launcher', 'main.ts'),
  outfile: path.join(outputDirectory, 'subminer.js'),
  target: 'bun',
  banner: '#!/usr/bin/env bun',
});
bundle({
  entrypoint: path.join(process.cwd(), 'src', 'main', 'runtime', 'prepare-launcher-runtime.ts'),
  outfile: path.join(outputDirectory, 'prepare.cjs'),
  target: 'node',
  format: 'cjs',
});

const posixLauncherPath = path.join(outputDirectory, 'subminer');
fs.writeFileSync(posixLauncherPath, posixLauncherBootstrapContent(), { mode: 0o755 });
fs.chmodSync(posixLauncherPath, 0o755);
fs.writeFileSync(
  path.join(outputDirectory, 'subminer.cmd'),
  windowsLauncherBootstrapContent(),
  'utf8',
);
fs.writeFileSync(path.join(outputDirectory, 'version'), `${packageJson.version}\n`, 'utf8');

console.log(`Built launcher runtime artifacts in ${outputDirectory}`);
