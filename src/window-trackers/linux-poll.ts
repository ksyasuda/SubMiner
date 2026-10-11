/*
  SubMiner - All-in-one sentence mining overlay
  Copyright (C) 2024 sudacode

  This program is free software: you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation, either version 3 of the License, or
  (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

import { execFile } from 'child_process';

// Shared polling helpers for the Hyprland and Sway trackers, which poll a
// compositor CLI on a short interval from the Electron main thread.

export type CommandRunner = (command: string, args: string[]) => Promise<string>;

const COMMAND_TIMEOUT_MS = 1500;

// Runs a command off the main thread. The timeout (SIGKILL, so a wedged client
// can't ignore it) keeps a hung compositor from stalling the poll loop forever.
export function execFileWithTimeout(command: string, args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(
      command,
      args,
      { encoding: 'utf-8', timeout: COMMAND_TIMEOUT_MS, killSignal: 'SIGKILL' },
      (error, stdout) => {
        if (error) {
          reject(error);
          return;
        }
        resolve(stdout);
      },
    );
  });
}

// Wraps an async task so overlapping calls coalesce: at most one run is in
// flight, and every call made during it shares a single follow-up run. The
// returned promise settles after a run that started no earlier than the call.
export function createCoalescedTask(task: () => Promise<void>): () => Promise<void> {
  let active: Promise<void> | null = null;
  let queued: Promise<void> | null = null;

  const run = (): Promise<void> => {
    const current: Promise<void> = task().finally(() => {
      if (active === current) {
        active = null;
      }
    });
    active = current;
    return current;
  };

  return () => {
    if (queued) {
      return queued;
    }
    if (!active) {
      return run();
    }
    const next = (): Promise<void> => {
      queued = null;
      return run();
    };
    queued = active.then(next, next);
    return queued;
  };
}

// Caches process command lines by pid. A pid's argv doesn't change, so each
// pid costs one `ps` spawn. Pids missing from the latest lookup are evicted,
// which keeps the cache bounded and re-reads a pid if it gets reused.
export class ProcessCommandLineCache {
  private readonly commandLines = new Map<number, string>();

  constructor(private readonly runCommand: CommandRunner) {}

  // A pid whose `ps` call fails (e.g. it exited mid-poll) is left out of the
  // result instead of failing the lookup. Empty results aren't cached, so that
  // pid is retried on the next poll.
  async lookup(pids: readonly number[]): Promise<ReadonlyMap<number, string>> {
    const wanted = new Set(pids);
    for (const pid of this.commandLines.keys()) {
      if (!wanted.has(pid)) {
        this.commandLines.delete(pid);
      }
    }

    await Promise.all(
      [...wanted]
        .filter((pid) => !this.commandLines.has(pid))
        .map(async (pid) => {
          const commandLine = (
            await this.runCommand('ps', ['-p', String(pid), '-o', 'args=']).catch(() => '')
          ).trim();
          if (commandLine) {
            this.commandLines.set(pid, commandLine);
          }
        }),
    );
    return this.commandLines;
  }
}
