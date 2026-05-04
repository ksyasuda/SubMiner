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

export interface HyprlandPlacementBounds {
  x: number;
  y: number;
  width: number;
  height: number;
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
  bounds?: HyprlandPlacementBounds | null,
): string[][] {
  if (!client.address) {
    return [];
  }

  const windowAddress = `address:${client.address}`;
  const dispatches: string[][] = [];
  if (client.floating !== true) {
    dispatches.push(['dispatch', 'setfloating', windowAddress]);
  }
  if (client.pinned === true) {
    dispatches.push(['dispatch', 'pin', windowAddress]);
  }
  const roundedBounds = roundPlacementBounds(bounds);
  if (roundedBounds) {
    dispatches.push([
      'dispatch',
      'movewindowpixel',
      `exact ${roundedBounds.x} ${roundedBounds.y},${windowAddress}`,
    ]);
    dispatches.push([
      'dispatch',
      'resizewindowpixel',
      `exact ${roundedBounds.width} ${roundedBounds.height},${windowAddress}`,
    ]);
    dispatches.push(['dispatch', 'setprop', `${windowAddress} rounding 0`]);
    dispatches.push(['dispatch', 'setprop', `${windowAddress} border_size 0`]);
    dispatches.push(['dispatch', 'setprop', `${windowAddress} no_shadow 1`]);
    dispatches.push(['dispatch', 'setprop', `${windowAddress} no_blur 1`]);
    dispatches.push(['dispatch', 'setprop', `${windowAddress} decorate 0`]);
  }
  return dispatches;
}

function roundPlacementBounds(
  bounds?: HyprlandPlacementBounds | null,
): HyprlandPlacementBounds | null {
  if (!bounds) {
    return null;
  }
  const rounded = {
    x: Math.round(bounds.x),
    y: Math.round(bounds.y),
    width: Math.round(bounds.width),
    height: Math.round(bounds.height),
  };
  return Number.isFinite(rounded.x) &&
    Number.isFinite(rounded.y) &&
    Number.isFinite(rounded.width) &&
    Number.isFinite(rounded.height) &&
    rounded.width > 0 &&
    rounded.height > 0
    ? rounded
    : null;
}

export function ensureHyprlandWindowFloatingByTitle(options: {
  title: string;
  bounds?: HyprlandPlacementBounds | null;
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

    const dispatches = buildHyprlandPlacementDispatches(client, options.bounds);
    for (const args of dispatches) {
      run('hyprctl', args, { stdio: 'ignore' });
    }
    return dispatches.length > 0;
  } catch {
    return false;
  }
}
