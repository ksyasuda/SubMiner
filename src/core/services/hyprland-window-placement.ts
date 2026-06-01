import { execFileSync } from 'node:child_process';

export interface HyprlandPlacementClient {
  address?: string;
  at?: [number, number];
  floating?: boolean;
  hidden?: boolean;
  initialTitle?: string;
  mapped?: boolean;
  pid?: number;
  pinned?: boolean;
  size?: [number, number];
  title?: string;
}

export interface HyprlandPlacementBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface HyprlandPlacementDispatchOptions {
  configProvider?: HyprlandConfigProvider;
  promote?: boolean;
}

type ExecFileSync = typeof execFileSync;
export type HyprlandConfigProvider = 'hyprlang' | 'lua';

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

function readHyprlandPlacementClients(run: ExecFileSync): HyprlandPlacementClient[] {
  return parseHyprlandClients(String(run('hyprctl', ['-j', 'clients'], { encoding: 'utf-8' })));
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
  options: HyprlandPlacementDispatchOptions = {},
): string[][] {
  if (!client.address) {
    return [];
  }

  const windowAddress = `address:${client.address}`;
  const configProvider = options.configProvider ?? 'hyprlang';
  const dispatches: string[][] = [];
  if (client.floating !== true) {
    dispatches.push(
      configProvider === 'lua'
        ? luaWindowDispatch('float', windowAddress, ['action = "on"'])
        : ['dispatch', 'setfloating', windowAddress],
    );
  }
  if (client.pinned === true) {
    dispatches.push(
      configProvider === 'lua'
        ? luaWindowDispatch('pin', windowAddress, ['action = "off"'])
        : ['dispatch', 'pin', windowAddress],
    );
  }
  const roundedBounds = roundPlacementBounds(bounds);
  if (roundedBounds) {
    if (configProvider === 'lua') {
      dispatches.push(
        luaWindowDispatch('resize', windowAddress, [
          `x = ${roundedBounds.width}`,
          `y = ${roundedBounds.height}`,
        ]),
      );
      dispatches.push(
        luaWindowDispatch('move', windowAddress, [
          `x = ${roundedBounds.x}`,
          `y = ${roundedBounds.y}`,
        ]),
      );
      dispatches.push(luaWindowSetProp(windowAddress, 'rounding', '0'));
      dispatches.push(luaWindowSetProp(windowAddress, 'border_size', '0'));
      dispatches.push(luaWindowSetProp(windowAddress, 'no_shadow', '1'));
      dispatches.push(luaWindowSetProp(windowAddress, 'no_blur', '1'));
      dispatches.push(luaWindowSetProp(windowAddress, 'decorate', '0'));
    } else {
      dispatches.push([
        'dispatch',
        'resizewindowpixel',
        `exact ${roundedBounds.width} ${roundedBounds.height},${windowAddress}`,
      ]);
      dispatches.push([
        'dispatch',
        'movewindowpixel',
        `exact ${roundedBounds.x} ${roundedBounds.y},${windowAddress}`,
      ]);
      dispatches.push(['dispatch', 'setprop', `${windowAddress} rounding 0`]);
      dispatches.push(['dispatch', 'setprop', `${windowAddress} border_size 0`]);
      dispatches.push(['dispatch', 'setprop', `${windowAddress} no_shadow 1`]);
      dispatches.push(['dispatch', 'setprop', `${windowAddress} no_blur 1`]);
      dispatches.push(['dispatch', 'setprop', `${windowAddress} decorate 0`]);
    }
  }
  if (options.promote !== false) {
    dispatches.push(
      configProvider === 'lua'
        ? luaWindowDispatch('alter_zorder', windowAddress, ['mode = "top"'])
        : ['dispatch', 'alterzorder', `top,${windowAddress}`],
    );
  }
  return dispatches;
}

function luaWindowDispatch(name: string, windowAddress: string, fields: string[]): string[] {
  return [
    'dispatch',
    `hl.dsp.window.${name}({ ${[...fields, `window = ${luaString(windowAddress)}`].join(', ')} })`,
  ];
}

function luaWindowSetProp(windowAddress: string, prop: string, value: string): string[] {
  return luaWindowDispatch('set_prop', windowAddress, [
    `prop = ${luaString(prop)}`,
    `value = ${luaString(value)}`,
  ]);
}

function luaString(value: string): string {
  return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
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

function isFiniteTuple(value: unknown): value is [number, number] {
  return (
    Array.isArray(value) &&
    value.length >= 2 &&
    typeof value[0] === 'number' &&
    typeof value[1] === 'number' &&
    Number.isFinite(value[0]) &&
    Number.isFinite(value[1])
  );
}

export function getHyprlandClientPlacementBounds(
  client: HyprlandPlacementClient,
): HyprlandPlacementBounds | null {
  if (!isFiniteTuple(client.at) || !isFiniteTuple(client.size)) {
    return null;
  }

  return roundPlacementBounds({
    x: client.at[0],
    y: client.at[1],
    width: client.size[0],
    height: client.size[1],
  });
}

export function hyprlandPlacementBoundsMatch(
  actual: HyprlandPlacementBounds | null,
  target: HyprlandPlacementBounds | null,
  tolerancePx = 1,
): boolean {
  const roundedActual = roundPlacementBounds(actual);
  const roundedTarget = roundPlacementBounds(target);
  if (!roundedActual || !roundedTarget) {
    return false;
  }

  return (
    Math.abs(roundedActual.x - roundedTarget.x) <= tolerancePx &&
    Math.abs(roundedActual.y - roundedTarget.y) <= tolerancePx &&
    Math.abs(roundedActual.width - roundedTarget.width) <= tolerancePx &&
    Math.abs(roundedActual.height - roundedTarget.height) <= tolerancePx
  );
}

function clientMatchesPlacementBounds(
  client: HyprlandPlacementClient,
  bounds: HyprlandPlacementBounds,
): boolean | null {
  const actual = getHyprlandClientPlacementBounds(client);
  return actual ? hyprlandPlacementBoundsMatch(actual, bounds) : null;
}

export function hasHyprlandWindowPlacementBoundsMismatch(options: {
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

  const targetBounds = roundPlacementBounds(options.bounds);
  if (!targetBounds) {
    return false;
  }

  const run = options.execFileSync ?? execFileSync;
  try {
    const client = findHyprlandWindowForPlacement(readHyprlandPlacementClients(run), {
      pid: options.pid ?? process.pid,
      title: options.title,
    });
    if (!client) {
      return false;
    }
    return clientMatchesPlacementBounds(client, targetBounds) === false;
  } catch {
    return false;
  }
}

export function ensureHyprlandWindowFloatingByTitle(options: {
  title: string;
  bounds?: HyprlandPlacementBounds | null;
  platform?: NodeJS.Platform;
  env?: NodeJS.ProcessEnv;
  pid?: number;
  promote?: boolean;
  execFileSync?: ExecFileSync;
}): boolean {
  if (!shouldAttemptHyprlandWindowPlacement(options.platform, options.env)) {
    return false;
  }

  const run = options.execFileSync ?? execFileSync;
  try {
    const clients = readHyprlandPlacementClients(run);
    const client = findHyprlandWindowForPlacement(clients, {
      pid: options.pid ?? process.pid,
      title: options.title,
    });
    if (!client) {
      return false;
    }

    const configProvider = detectHyprlandConfigProvider(run);
    const targetBounds = roundPlacementBounds(options.bounds);
    const shouldVerifyBounds =
      targetBounds !== null && clientMatchesPlacementBounds(client, targetBounds) === false;
    const dispatches = buildHyprlandPlacementDispatches(client, options.bounds, {
      configProvider,
      promote: options.promote,
    });
    for (const args of dispatches) {
      run('hyprctl', args, { stdio: 'ignore' });
    }
    if (shouldVerifyBounds) {
      try {
        const refreshedClient = findHyprlandWindowForPlacement(readHyprlandPlacementClients(run), {
          pid: options.pid ?? process.pid,
          title: options.title,
        });
        if (
          refreshedClient &&
          targetBounds &&
          clientMatchesPlacementBounds(refreshedClient, targetBounds) === false
        ) {
          for (const args of buildHyprlandPlacementDispatches(refreshedClient, targetBounds, {
            configProvider,
            promote: options.promote,
          })) {
            run('hyprctl', args, { stdio: 'ignore' });
          }
        }
      } catch {
        // Best-effort reconciliation: the initial placement dispatches already ran.
      }
    }
    return dispatches.length > 0;
  } catch {
    return false;
  }
}

function detectHyprlandConfigProvider(run: ExecFileSync): HyprlandConfigProvider {
  try {
    return parseHyprlandConfigProvider(
      String(run('hyprctl', ['-j', 'status'], { encoding: 'utf-8' })),
    );
  } catch {
    return 'hyprlang';
  }
}

function parseHyprlandConfigProvider(output: string): HyprlandConfigProvider {
  const payloadStart = output.indexOf('{');
  if (payloadStart < 0) {
    return 'hyprlang';
  }

  const parsed = JSON.parse(output.slice(payloadStart)) as unknown;
  return isHyprlandStatusPayload(parsed) && parsed.configProvider === 'lua' ? 'lua' : 'hyprlang';
}

function isHyprlandStatusPayload(value: unknown): value is { configProvider?: unknown } {
  return Boolean(value) && typeof value === 'object';
}
