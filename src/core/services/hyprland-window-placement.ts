import { execFileSync } from 'node:child_process';

export interface HyprlandPlacementClient {
  address?: string;
  floating?: boolean;
  hidden?: boolean;
  initialTitle?: string;
  mapped?: boolean;
  pid?: number;
  pinned?: boolean;
  title?: string;
}

type ExecFileSync = typeof execFileSync;

export function shouldAttemptHyprlandWindowPlacement(
  platform: NodeJS.Platform = process.platform,
  env: NodeJS.ProcessEnv = process.env,
): boolean {
  return platform === 'linux' && Boolean(env.HYPRLAND_INSTANCE_SIGNATURE);
}

function parseHyprlandClients(output: string): HyprlandPlacementClient[] {
  const payloadStart = output.indexOf('[');
  if (payloadStart < 0) {
    return [];
  }

  const parsed = JSON.parse(output.slice(payloadStart)) as unknown;
  return Array.isArray(parsed) ? (parsed as HyprlandPlacementClient[]) : [];
}

export function findHyprlandWindowForPlacement(
  clients: HyprlandPlacementClient[],
  options: {
    pid: number;
    title: string;
  },
): HyprlandPlacementClient | null {
  const title = options.title.trim();
  if (!title) {
    return null;
  }

  return (
    clients.find(
      (client) =>
        client.pid === options.pid &&
        client.address &&
        client.mapped !== false &&
        client.hidden !== true &&
        (client.title === title || client.initialTitle === title),
    ) ?? null
  );
}

export function buildHyprlandPlacementDispatches(
  client: HyprlandPlacementClient,
): string[][] {
  if (!client.address) {
    return [];
  }

  const windowAddress = `address:${client.address}`;
  const dispatches: string[][] = [];
  if (client.floating !== true) {
    dispatches.push(['dispatch', 'setfloating', windowAddress]);
  }
  if (client.pinned !== true) {
    dispatches.push(['dispatch', 'pin', windowAddress]);
  }
  return dispatches;
}

export function ensureHyprlandWindowFloatingByTitle(options: {
  title: string;
  platform?: NodeJS.Platform;
  env?: NodeJS.ProcessEnv;
  pid?: number;
  execFileSync?: ExecFileSync;
}): boolean {
  if (!shouldAttemptHyprlandWindowPlacement(options.platform, options.env)) {
    return false;
  }

  const run = options.execFileSync ?? execFileSync;
  try {
    const clients = parseHyprlandClients(
      String(run('hyprctl', ['-j', 'clients'], { encoding: 'utf-8' })),
    );
    const client = findHyprlandWindowForPlacement(clients, {
      pid: options.pid ?? process.pid,
      title: options.title,
    });
    if (!client) {
      return false;
    }

    const dispatches = buildHyprlandPlacementDispatches(client);
    for (const args of dispatches) {
      run('hyprctl', args, { stdio: 'ignore' });
    }
    return dispatches.length > 0;
  } catch {
    return false;
  }
}
