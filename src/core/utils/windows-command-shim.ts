import fs from 'node:fs';
import path from 'node:path';

/** True for batch shims (`.cmd`/`.bat`), which Windows only runs through cmd.exe. */
export function isWindowsCommandShim(command: string): boolean {
  return /\.(cmd|bat)$/i.test(command);
}

/**
 * Resolves an extensionless Windows command the way cmd.exe does: each `PATHEXT` extension in
 * the command's own directory, or in each `PATH` entry for a bare name. Returns `command`
 * unchanged when it already has an extension or nothing matches.
 */
export function resolveWindowsCommand(
  command: string,
  env: NodeJS.ProcessEnv,
  existsSync: (candidate: string) => boolean = fs.existsSync,
): string {
  if (path.extname(command)) return command;
  const directory = path.dirname(command);
  const directories =
    directory === '.' ? (env.PATH ?? env.Path ?? '').split(path.delimiter).filter(Boolean) : [''];
  const extensions = (env.PATHEXT ?? '.COM;.EXE;.BAT;.CMD').split(';').filter(Boolean);
  for (const entry of directories) {
    for (const extension of extensions) {
      const candidate = entry ? path.join(entry, command + extension) : command + extension;
      if (existsSync(candidate)) return candidate;
    }
  }
  return command;
}

/*!
 * Windows command escaping adapted from cross-spawn 7.0.6.
 *
 * The MIT License (MIT)
 *
 * Copyright (c) 2018 Made With MOXY Lda <hello@moxy.studio>
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
const WINDOWS_SHELL_META_CHARACTERS = /([()\][%!^"`<>&|;, *?])/g;

// Quote for both cmd.exe and the Windows argv parser. The outer caret escapes
// are consumed by cmd, leaving the quoted argument unchanged for the command.
function escapeWindowsShellArgument(value: string): string {
  const quotesEscaped = value
    .replace(/(?=(\\+?)?)\1"/g, '$1$1\\"')
    .replace(/(?=(\\+?)?)\1$/, '$1$1');
  return `"${quotesEscaped}"`.replace(WINDOWS_SHELL_META_CHARACTERS, '^$1');
}

/**
 * Builds the cmd.exe invocation that runs a batch shim with `args`. Spawn the result with
 * `windowsVerbatimArguments: true` so Node does not re-quote the prepared command line.
 */
export function buildWindowsShimInvocation(
  command: string,
  args: readonly string[],
  env: NodeJS.ProcessEnv,
): { file: string; args: string[] } {
  const commandLine = [command, ...args].map(escapeWindowsShellArgument).join(' ');
  return {
    file: env.ComSpec ?? env.COMSPEC ?? process.env.ComSpec ?? 'cmd.exe',
    args: ['/d', '/s', '/v:off', '/c', `"${commandLine}"`],
  };
}
