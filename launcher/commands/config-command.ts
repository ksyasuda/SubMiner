import fs from 'node:fs';
import { fail } from '../log.js';
import { resolveMainConfigPath } from '../config-path.js';
import type { LauncherCommandContext } from './context.js';

interface ConfigCommandDeps {
  existsSync(path: string): boolean;
  readFileSync(path: string, encoding: BufferEncoding): string;
  resolveMainConfigPath(): string;
}

const defaultDeps: ConfigCommandDeps = {
  existsSync: fs.existsSync,
  readFileSync: fs.readFileSync,
  resolveMainConfigPath,
};

export function runConfigCommand(
  context: LauncherCommandContext,
  deps: ConfigCommandDeps = defaultDeps,
): boolean {
  const { args, processAdapter } = context;
  if (args.configPath) {
    processAdapter.writeStdout(`${deps.resolveMainConfigPath()}\n`);
    return true;
  }

  if (!args.configShow) {
    return false;
  }

  const configPath = deps.resolveMainConfigPath();
  if (!deps.existsSync(configPath)) {
    fail(`Config file not found: ${configPath}`);
  }

  const contents = deps.readFileSync(configPath, 'utf8');
  processAdapter.writeStdout(contents);
  if (!contents.endsWith('\n')) {
    processAdapter.writeStdout('\n');
  }
  return true;
}
