import * as fs from 'fs';
import * as path from 'path';
import { DEFAULT_CONFIG, deepCloneConfig, generateConfigTemplate } from './config';
import { getDefaultMpvSocketPath } from './shared/mpv-socket-path';

const CONFIG_EXAMPLE_PLATFORM: NodeJS.Platform = 'win32';

export function generateConfigExampleTemplate(): string {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  config.mpv.socketPath = getDefaultMpvSocketPath(CONFIG_EXAMPLE_PLATFORM);
  return generateConfigTemplate(config);
}

type ConfigExampleFsDeps = {
  existsSync?: (candidate: string) => boolean;
  mkdirSync?: (candidate: string, options: { recursive: true }) => void;
  writeFileSync?: (candidate: string, content: string, encoding: BufferEncoding) => void;
  log?: (message: string) => void;
};

export function resolveConfigExampleOutputPaths(options?: {
  cwd?: string;
  docsSiteDirName?: string;
  existsSync?: (candidate: string) => boolean;
}): string[] {
  const cwd = options?.cwd ?? process.cwd();
  const existsSync = options?.existsSync ?? fs.existsSync;
  const docsSiteDirName = options?.docsSiteDirName ?? 'docs-site';
  const outputPaths = [path.join(cwd, 'config.example.jsonc')];
  const docsSiteRoot = path.join(cwd, docsSiteDirName);

  if (existsSync(docsSiteRoot)) {
    outputPaths.push(path.join(docsSiteRoot, 'public', 'config.example.jsonc'));
  }

  return outputPaths;
}

export function writeConfigExampleArtifacts(
  template: string,
  options?: {
    cwd?: string;
    docsSiteDirName?: string;
    deps?: ConfigExampleFsDeps;
  },
): string[] {
  const mkdirSync = options?.deps?.mkdirSync ?? fs.mkdirSync;
  const writeFileSync = options?.deps?.writeFileSync ?? fs.writeFileSync;
  const log = options?.deps?.log ?? console.log;
  const outputPaths = resolveConfigExampleOutputPaths({
    cwd: options?.cwd,
    docsSiteDirName: options?.docsSiteDirName,
    existsSync: options?.deps?.existsSync,
  });

  for (const outputPath of outputPaths) {
    mkdirSync(path.dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, template, 'utf-8');
    log(`Generated ${outputPath}`);
  }

  return outputPaths;
}

function main(): void {
  const template = generateConfigExampleTemplate();
  writeConfigExampleArtifacts(template);
}

if (require.main === module) {
  main();
}
