import type { ChildProcess } from 'node:child_process';
import { once } from 'node:events';

export type EnvDelta = { set: Record<string, string>; unset: string[] };

export function applyEnvDelta(base: NodeJS.ProcessEnv, ...deltas: EnvDelta[]): NodeJS.ProcessEnv {
  const env = { ...base };
  for (const delta of deltas) {
    for (const name of delta.unset) delete env[name];
    Object.assign(env, delta.set);
  }
  return env;
}

/** Terminates a child and waits for it to exit, escalating to SIGKILL after `graceMs`. */
export async function stopProcess(child: ChildProcess, graceMs = 5_000): Promise<void> {
  // A child that never spawned (missing binary) has no pid and never emits exit.
  if (child.pid === undefined || child.exitCode !== null || child.signalCode !== null) return;
  const exited = once(child, 'exit');
  child.kill();
  const timer = setTimeout(() => child.kill('SIGKILL'), graceMs);
  await exited;
  clearTimeout(timer);
}
