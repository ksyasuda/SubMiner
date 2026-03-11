import * as fs from 'fs';
import * as path from 'path';
import { DEFAULT_CONFIG, generateConfigTemplate } from './config';

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
  const template = generateConfigTemplate(DEFAULT_CONFIG);
  writeConfigExampleArtifacts(template);
}

if (require.main === module) {
  main();
}
