import { execFile } from 'node:child_process';
import os from 'node:os';
import { isAbsolute } from 'node:path';

const MARKER = '__SUBMINER_LOGIN_PATH__';
const DEFAULT_TIMEOUT_MS = 5000;

export type RunShell = (shell: string, args: string[], timeoutMs: number) => Promise<string>;

type LoginShellPathOptions = {
  env: NodeJS.ProcessEnv;
  shell?: string;
  runShell?: RunShell;
  timeoutMs?: number;
};

const runShellDefault: RunShell = (shell, args, timeoutMs) =>
  new Promise((resolve, reject) => {
    const child = execFile(
      shell,
      args,
      {
        timeout: timeoutMs,
        encoding: 'utf8',
        env: { ...process.env, DISABLE_AUTO_UPDATE: 'true' },
      },
      (error, stdout) => (error ? reject(error) : resolve(stdout)),
    );
    child.stdin?.end();
  });

function defaultShell(env: NodeJS.ProcessEnv): string {
  if (env.SHELL && isAbsolute(env.SHELL)) return env.SHELL;
  try {
    const shell = os.userInfo().shell;
    return shell && isAbsolute(shell) ? shell : '/bin/zsh';
  } catch {
    return '/bin/zsh';
  }
}

/**
 * Reads PATH as the user's interactive login shell sees it. macOS GUI apps inherit
 * launchd's minimal PATH, so dirs added in ~/.zshrc and friends are otherwise invisible.
 */
export async function readLoginShellPath(options: LoginShellPathOptions): Promise<string | null> {
  const shell = options.shell ?? defaultShell(options.env);
  if (!isAbsolute(shell)) throw new Error('Login shell must be an absolute path');
  const run = options.runShell ?? runShellDefault;
  // printenv keeps this shell-agnostic (fish exposes $PATH as a list).
  const command = `printf '%s' '${MARKER}'; /usr/bin/printenv PATH; printf '%s' '${MARKER}'`;
  const stdout = await run(shell, ['-ilc', command], options.timeoutMs ?? DEFAULT_TIMEOUT_MS);
  const value = stdout.split(MARKER)[1]?.trim();
  return value || null;
}

/** Puts login-shell entries first (terminal precedence), keeping any process-only entries after. */
export function mergePathValues(loginPath: string, currentPath: string | undefined): string {
  const merged: string[] = [];
  for (const entry of [...loginPath.split(':'), ...(currentPath ?? '').split(':')]) {
    if (entry && !merged.includes(entry)) merged.push(entry);
  }
  return merged.join(':');
}

/** Merges the login-shell PATH into `options.env.PATH`. Resolves false when it couldn't be read. */
export async function applyLoginShellPath(options: LoginShellPathOptions): Promise<boolean> {
  const loginPath = await readLoginShellPath(options);
  if (!loginPath) return false;
  options.env.PATH = mergePathValues(loginPath, options.env.PATH);
  return true;
}
