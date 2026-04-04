import { Client } from '@xhayper/discord-rpc';

import type { DiscordActivityPayload } from '../../core/services/discord-presence';

type DiscordRpcClientUserLike = {
  setActivity: (activity: DiscordActivityPayload) => Promise<unknown>;
  clearActivity: () => Promise<void>;
};

type DiscordRpcRawClient = {
  login: () => Promise<void>;
  destroy: () => Promise<void>;
  user?: DiscordRpcClientUserLike;
};

export type DiscordRpcClient = {
  login: () => Promise<void>;
  setActivity: (activity: DiscordActivityPayload) => Promise<void>;
  clearActivity: () => Promise<void>;
  destroy: () => Promise<void>;
};

function requireUser(client: DiscordRpcRawClient): DiscordRpcClientUserLike {
  if (!client.user) {
    throw new Error('Discord RPC client user is unavailable');
  }

  return client.user;
}

export function wrapDiscordRpcClient(client: DiscordRpcRawClient): DiscordRpcClient {
  return {
    login: () => client.login(),
    setActivity: (activity) =>
      requireUser(client)
        .setActivity(activity)
        .then(() => undefined),
    clearActivity: () => requireUser(client).clearActivity(),
    destroy: () => client.destroy(),
  };
}

export function createDiscordRpcClient(
  clientId: string,
  deps?: {
    createClient?: (options: {
      clientId: string;
      transport: { type: 'ipc' };
    }) => DiscordRpcRawClient;
  },
): DiscordRpcClient {
  const client =
    deps?.createClient?.({ clientId, transport: { type: 'ipc' } }) ??
    new Client({ clientId, transport: { type: 'ipc' } });

  return wrapDiscordRpcClient(client);
}
