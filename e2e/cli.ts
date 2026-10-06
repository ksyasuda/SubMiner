// Keeps one isolated e2e session running so a change can be poked at by hand
// (or by an agent) without touching the real desktop:
//
//   bun run e2e start                      boot and stay up until `stop` or Ctrl-C
//   bun run e2e eval <page> <expression>   evaluate JS in a window (or "main" for the main process)
//   bun run e2e shot <page> <file.png>     save a screenshot of a window
//   bun run e2e mpv <command> [args...]    send an mpv IPC command (args parsed as JSON)
//   bun run e2e app <args...>              run a SubMiner CLI command against the instance
//   bun run e2e anki                       print the AnkiConnect requests received so far
//   bun run e2e stop
//
// <page> is a URL fragment: "overlay" is shorthand for the overlay window.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { connectCdp, connectCdpPage, listCdpTargets } from './harness/cdp';
import { connectMpv } from './harness/mpv';
import { OVERLAY_PAGE, runAppCommand, startE2eSession, type AppTarget } from './harness/session';
import { waitUntil } from './harness/wait';

type SessionState = {
  pid: number;
  root: string;
  display: string;
  cdpPort: number;
  inspectorUrl: string;
  mpvSocketPath: string;
  ankiUrl: string;
  target: AppTarget;
};

const stateFile = path.join(os.tmpdir(), 'subminer-e2e-session.json');

function readState(): SessionState {
  if (!fs.existsSync(stateFile)) {
    throw new Error('No e2e session is running. Start one with `bun run e2e start`.');
  }
  return JSON.parse(fs.readFileSync(stateFile, 'utf8')) as SessionState;
}

function isAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function parseJsonOrString(value: string): unknown {
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

async function connectNamedPage(state: SessionState, pageName: string) {
  if (pageName === 'main') return connectCdp(state.inspectorUrl, 'main process');
  const urlPart = pageName === 'overlay' ? OVERLAY_PAGE : pageName;
  const targets = await listCdpTargets(state.cdpPort);
  const target = targets.find((candidate) => candidate.url.includes(urlPart));
  if (!target) {
    const open = targets.map((candidate) => `  ${candidate.type} ${candidate.url.slice(0, 100)}`);
    throw new Error(`No window matches "${pageName}". Open targets:\n${open.join('\n')}`);
  }
  return connectCdpPage(target);
}

async function withPage<T>(
  state: SessionState,
  pageName: string,
  use: (page: Awaited<ReturnType<typeof connectCdpPage>>) => Promise<T>,
): Promise<T> {
  const page = await connectNamedPage(state, pageName);
  try {
    return await use(page);
  } finally {
    page.close();
  }
}

async function start(): Promise<void> {
  if (fs.existsSync(stateFile) && isAlive(readState().pid)) {
    throw new Error('An e2e session is already running. Stop it with `bun run e2e stop`.');
  }
  const session = await startE2eSession();
  const state: SessionState = {
    pid: process.pid,
    root: session.root,
    display: session.display,
    cdpPort: session.cdpPort,
    inspectorUrl: session.inspectorUrl,
    mpvSocketPath: session.mpvSocketPath,
    ankiUrl: session.anki.url,
    target: session.target,
  };
  // The state file grants main-process code evaluation via inspectorUrl, so keep it
  // owner-only (mode applies on create, hence the rm of any stale file first).
  fs.rmSync(stateFile, { force: true });
  fs.writeFileSync(stateFile, JSON.stringify(state, null, 2), { mode: 0o600 });
  console.log(JSON.stringify(state, null, 2));

  await new Promise<void>((resolve) => {
    process.once('SIGINT', resolve);
    process.once('SIGTERM', resolve);
  });
  fs.rmSync(stateFile, { force: true });
  await session.dispose();
}

async function main(): Promise<void> {
  const [command, ...args] = process.argv.slice(2);
  if (command === 'start') return start();

  const state = readState();
  switch (command) {
    case 'eval': {
      const [pageName, expression] = args;
      if (!pageName || !expression) throw new Error('Usage: eval <page> <expression>');
      const value = await withPage(state, pageName, (page) => page.evaluate<unknown>(expression));
      console.log(JSON.stringify(value, null, 2));
      return;
    }
    case 'shot': {
      const [pageName, outputPath] = args;
      if (!pageName || !outputPath) throw new Error('Usage: shot <page> <file.png>');
      fs.writeFileSync(outputPath, await withPage(state, pageName, (page) => page.screenshot()));
      return;
    }
    case 'mpv': {
      if (args.length === 0) throw new Error('Usage: mpv <command> [args...]');
      const mpv = await connectMpv(state.mpvSocketPath);
      try {
        console.log(JSON.stringify((await mpv.command(...args.map(parseJsonOrString))) ?? null));
      } finally {
        mpv.close();
      }
      return;
    }
    case 'app':
      await runAppCommand(state.target, args);
      return;
    case 'anki':
      console.log(JSON.stringify(await (await fetch(state.ankiUrl)).json(), null, 2));
      return;
    case 'stop':
      process.kill(state.pid, 'SIGTERM');
      await waitUntil(() => !isAlive(state.pid), { description: 'e2e session to exit' });
      return;
    default:
      throw new Error(`Unknown command "${command ?? ''}". See the header of e2e/cli.ts.`);
  }
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
