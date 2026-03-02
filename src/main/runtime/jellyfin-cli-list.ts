import type { CliArgs } from '../../cli/args';

type JellyfinSession = {
  serverUrl: string;
  accessToken: string;
  userId: string;
  username: string;
};

type JellyfinClientInfo = {
  clientName: string;
  clientVersion: string;
  deviceId: string;
};

type JellyfinConfig = {
  defaultLibraryId: string;
};

type JellyfinPreviewAuthPayload = {
  serverUrl: string;
  accessToken: string;
  userId: string;
};

export function createHandleJellyfinListCommands(deps: {
  listJellyfinLibraries: (
    session: JellyfinSession,
    clientInfo: JellyfinClientInfo,
  ) => Promise<Array<{ id: string; name: string; collectionType?: string; type?: string }>>;
  listJellyfinItems: (
    session: JellyfinSession,
    clientInfo: JellyfinClientInfo,
    params: {
      libraryId: string;
      searchTerm?: string;
      limit: number;
      recursive?: boolean;
      includeItemTypes?: string;
    },
  ) => Promise<Array<{ id: string; title: string; type: string }>>;
  listJellyfinSubtitleTracks: (
    session: JellyfinSession,
    clientInfo: JellyfinClientInfo,
    itemId: string,
  ) => Promise<
    Array<{
      index: number;
      language?: string;
      title?: string;
      deliveryMethod?: string;
      codec?: string;
      isDefault?: boolean;
      isForced?: boolean;
      isExternal?: boolean;
      deliveryUrl?: string | null;
    }>
  >;
  writeJellyfinPreviewAuth: (responsePath: string, payload: JellyfinPreviewAuthPayload) => void;
  logInfo: (message: string) => void;
}) {
  return async (params: {
    args: CliArgs;
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: JellyfinConfig;
  }): Promise<boolean> => {
    const { args, session, clientInfo, jellyfinConfig } = params;

    if (args.jellyfinPreviewAuth) {
      const responsePath = args.jellyfinResponsePath?.trim();
      if (!responsePath) {
        throw new Error('Missing --jellyfin-response-path for --jellyfin-preview-auth.');
      }
      deps.writeJellyfinPreviewAuth(responsePath, {
        serverUrl: session.serverUrl,
        accessToken: session.accessToken,
        userId: session.userId,
      });
      deps.logInfo('Jellyfin preview auth written.');
      return true;
    }

    if (args.jellyfinLibraries) {
      const libraries = await deps.listJellyfinLibraries(session, clientInfo);
      if (libraries.length === 0) {
        deps.logInfo('No Jellyfin libraries found.');
        return true;
      }
      for (const library of libraries) {
        deps.logInfo(
          `Jellyfin library: ${library.name} [${library.id}] (${library.collectionType || library.type || 'unknown'})`,
        );
      }
      return true;
    }

    if (args.jellyfinItems) {
      const libraryId = args.jellyfinLibraryId || jellyfinConfig.defaultLibraryId;
      if (!libraryId) {
        throw new Error(
          'Missing Jellyfin library id. Use --jellyfin-library-id or set jellyfin.defaultLibraryId.',
        );
      }
      const items = await deps.listJellyfinItems(session, clientInfo, {
        libraryId,
        searchTerm: args.jellyfinSearch,
        limit: args.jellyfinLimit ?? 100,
        recursive: args.jellyfinRecursive,
        includeItemTypes: args.jellyfinIncludeItemTypes,
      });
      if (items.length === 0) {
        deps.logInfo('No Jellyfin items found for the selected library/search.');
        return true;
      }
      for (const item of items) {
        deps.logInfo(`Jellyfin item: ${item.title} [${item.id}] (${item.type})`);
      }
      return true;
    }

    if (args.jellyfinSubtitles) {
      if (!args.jellyfinItemId) {
        throw new Error('Missing --jellyfin-item-id for --jellyfin-subtitles.');
      }
      const tracks = await deps.listJellyfinSubtitleTracks(
        session,
        clientInfo,
        args.jellyfinItemId,
      );
      if (tracks.length === 0) {
        deps.logInfo('No Jellyfin subtitle tracks found for item.');
        return true;
      }
      for (const track of tracks) {
        if (args.jellyfinSubtitleUrlsOnly) {
          if (track.deliveryUrl) deps.logInfo(track.deliveryUrl);
          continue;
        }
        deps.logInfo(
          `Jellyfin subtitle: index=${track.index} lang=${track.language || 'unknown'} title="${track.title || '-'}" method=${track.deliveryMethod || 'unknown'} codec=${track.codec || 'unknown'} default=${track.isDefault ? 'yes' : 'no'} forced=${track.isForced ? 'yes' : 'no'} external=${track.isExternal ? 'yes' : 'no'} url=${track.deliveryUrl || '-'}`,
        );
      }
      return true;
    }

    return false;
  };
}
