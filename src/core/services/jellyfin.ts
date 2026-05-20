import { JellyfinConfig } from '../../types';

const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_LOGIN_TIMEOUT_MS = 15_000;

export interface JellyfinAuthSession {
  serverUrl: string;
  accessToken: string;
  userId: string;
  username: string;
}

export interface JellyfinLibrary {
  id: string;
  name: string;
  collectionType: string;
  type: string;
}

export interface JellyfinPlaybackSelection {
  itemId: string;
  audioStreamIndex?: number;
  subtitleStreamIndex?: number;
}

export interface JellyfinPlaybackPlan {
  mode: 'direct' | 'transcode';
  url: string;
  title: string;
  startTimeTicks: number;
  audioStreamIndex: number | null;
  subtitleStreamIndex: number | null;
}

export interface JellyfinSubtitleTrack {
  index: number;
  language: string;
  title: string;
  codec: string;
  isDefault: boolean;
  isForced: boolean;
  isExternal: boolean;
  deliveryMethod: string;
  deliveryUrl: string | null;
}

interface JellyfinAuthResponse {
  AccessToken?: string;
  User?: { Id?: string; Name?: string };
}

interface JellyfinMediaStream {
  Index?: number;
  Type?: string;
  IsExternal?: boolean;
  IsDefault?: boolean;
  IsForced?: boolean;
  Language?: string;
  DisplayTitle?: string;
  Title?: string;
  Codec?: string;
  DeliveryMethod?: string;
  DeliveryUrl?: string;
  IsExternalUrl?: boolean;
}

interface JellyfinMediaSource {
  Id?: string;
  Container?: string;
  SupportsDirectStream?: boolean;
  SupportsTranscoding?: boolean;
  TranscodingUrl?: string;
  DefaultAudioStreamIndex?: number;
  DefaultSubtitleStreamIndex?: number;
  MediaStreams?: JellyfinMediaStream[];
  LiveStreamId?: string;
}

interface JellyfinItemUserData {
  PlaybackPositionTicks?: number;
}

interface JellyfinItem {
  Id?: string;
  Name?: string;
  Type?: string;
  SeriesName?: string;
  ParentIndexNumber?: number;
  IndexNumber?: number;
  UserData?: JellyfinItemUserData;
  MediaSources?: JellyfinMediaSource[];
}

interface JellyfinItemsResponse {
  Items?: JellyfinItem[];
}

interface JellyfinPlaybackInfoResponse {
  MediaSources?: JellyfinMediaSource[];
}

export interface JellyfinClientInfo {
  deviceId: string;
  clientName: string;
  clientVersion: string;
}

function normalizeBaseUrl(value: string): string {
  return value.trim().replace(/\/+$/, '');
}

function ensureString(value: unknown, fallback = ''): string {
  return typeof value === 'string' ? value : fallback;
}

function asIntegerOrNull(value: unknown): number | null {
  return typeof value === 'number' && Number.isInteger(value) ? value : null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isAbortError(error: unknown): boolean {
  return isRecord(error) && error.name === 'AbortError';
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }
  return String(error || 'unknown error');
}

function resolveDeliveryUrl(
  session: JellyfinAuthSession,
  stream: JellyfinMediaStream,
  itemId: string,
  mediaSourceId: string,
): string | null {
  const deliveryUrl = ensureString(stream.DeliveryUrl).trim();
  if (deliveryUrl) {
    if (stream.IsExternalUrl === true) return deliveryUrl;
    const resolved = new URL(deliveryUrl, `${session.serverUrl}/`);
    if (!resolved.searchParams.has('api_key')) {
      resolved.searchParams.set('api_key', session.accessToken);
    }
    return resolved.toString();
  }

  const streamIndex = asIntegerOrNull(stream.Index);
  if (streamIndex === null || !itemId || !mediaSourceId) return null;
  const codec = ensureString(stream.Codec).toLowerCase();
  const ext =
    codec === 'subrip'
      ? 'srt'
      : codec === 'webvtt'
        ? 'vtt'
        : codec === 'vtt'
          ? 'vtt'
          : codec === 'ass'
            ? 'ass'
            : codec === 'ssa'
              ? 'ssa'
              : 'srt';
  const fallback = new URL(
    `/Videos/${encodeURIComponent(itemId)}/${encodeURIComponent(mediaSourceId)}/Subtitles/${streamIndex}/Stream.${ext}`,
    `${session.serverUrl}/`,
  );
  if (!fallback.searchParams.has('api_key')) {
    fallback.searchParams.set('api_key', session.accessToken);
  }
  return fallback.toString();
}

function createAuthorizationHeader(client: JellyfinClientInfo, token?: string): string {
  const parts = [
    `Client="${client.clientName}"`,
    `Device="${client.clientName}"`,
    `DeviceId="${client.deviceId}"`,
    `Version="${client.clientVersion}"`,
  ];
  if (token) parts.push(`Token="${token}"`);
  return `MediaBrowser ${parts.join(', ')}`;
}

async function jellyfinRequestJson<T>(
  path: string,
  init: RequestInit,
  session: JellyfinAuthSession,
  client: JellyfinClientInfo,
): Promise<T> {
  const headers = new Headers(init.headers ?? {});
  headers.set('Content-Type', 'application/json');
  headers.set('Authorization', createAuthorizationHeader(client, session.accessToken));
  headers.set('X-Emby-Token', session.accessToken);

  const response = await fetch(`${session.serverUrl}${path}`, {
    ...init,
    headers,
  });

  if (response.status === 401 || response.status === 403) {
    throw new Error('Jellyfin authentication failed (invalid or expired token).');
  }
  if (!response.ok) {
    throw new Error(`Jellyfin request failed (${response.status} ${response.statusText}).`);
  }
  return response.json() as Promise<T>;
}

function createDirectPlayUrl(
  session: JellyfinAuthSession,
  itemId: string,
  mediaSource: JellyfinMediaSource,
  plan: JellyfinPlaybackPlan,
): string {
  const query = new URLSearchParams({
    static: 'true',
    api_key: session.accessToken,
    MediaSourceId: ensureString(mediaSource.Id),
  });
  if (mediaSource.LiveStreamId) {
    query.set('LiveStreamId', mediaSource.LiveStreamId);
  }
  if (plan.audioStreamIndex !== null) {
    query.set('AudioStreamIndex', String(plan.audioStreamIndex));
  }
  if (plan.subtitleStreamIndex !== null) {
    query.set('SubtitleStreamIndex', String(plan.subtitleStreamIndex));
  }
  if (plan.startTimeTicks > 0) {
    query.set('StartTimeTicks', String(plan.startTimeTicks));
  }
  return `${session.serverUrl}/Videos/${itemId}/stream?${query.toString()}`;
}

function createTranscodeUrl(
  session: JellyfinAuthSession,
  itemId: string,
  mediaSource: JellyfinMediaSource,
  plan: JellyfinPlaybackPlan,
  config: JellyfinConfig,
): string {
  if (mediaSource.TranscodingUrl) {
    const url = new URL(`${session.serverUrl}${mediaSource.TranscodingUrl}`);
    if (!url.searchParams.has('api_key')) {
      url.searchParams.set('api_key', session.accessToken);
    }
    if (!url.searchParams.has('AudioStreamIndex') && plan.audioStreamIndex !== null) {
      url.searchParams.set('AudioStreamIndex', String(plan.audioStreamIndex));
    }
    if (!url.searchParams.has('SubtitleStreamIndex') && plan.subtitleStreamIndex !== null) {
      url.searchParams.set('SubtitleStreamIndex', String(plan.subtitleStreamIndex));
    }
    if (!url.searchParams.has('StartTimeTicks') && plan.startTimeTicks > 0) {
      url.searchParams.set('StartTimeTicks', String(plan.startTimeTicks));
    }
    return url.toString();
  }

  const query = new URLSearchParams({
    api_key: session.accessToken,
    MediaSourceId: ensureString(mediaSource.Id),
    VideoCodec: ensureString(config.transcodeVideoCodec, 'h264'),
    TranscodingContainer: 'ts',
  });
  if (plan.audioStreamIndex !== null) {
    query.set('AudioStreamIndex', String(plan.audioStreamIndex));
  }
  if (plan.subtitleStreamIndex !== null) {
    query.set('SubtitleStreamIndex', String(plan.subtitleStreamIndex));
  }
  if (plan.startTimeTicks > 0) {
    query.set('StartTimeTicks', String(plan.startTimeTicks));
  }
  return `${session.serverUrl}/Videos/${itemId}/master.m3u8?${query.toString()}`;
}

function getStreamDefaults(source: JellyfinMediaSource): {
  audioStreamIndex: number | null;
} {
  const audioDefault = asIntegerOrNull(source.DefaultAudioStreamIndex);
  if (audioDefault !== null) return { audioStreamIndex: audioDefault };

  const streams = Array.isArray(source.MediaStreams) ? source.MediaStreams : [];
  const defaultAudio = streams.find(
    (stream) => stream.Type === 'Audio' && stream.IsDefault === true,
  );
  return {
    audioStreamIndex: asIntegerOrNull(defaultAudio?.Index),
  };
}

function getDisplayTitle(item: JellyfinItem): string {
  if (item.Type === 'Episode') {
    const season = asIntegerOrNull(item.ParentIndexNumber) ?? 0;
    const episode = asIntegerOrNull(item.IndexNumber) ?? 0;
    const prefix = item.SeriesName ? `${item.SeriesName} ` : '';
    return `${prefix}S${String(season).padStart(2, '0')}E${String(episode).padStart(2, '0')} ${ensureString(item.Name).trim()}`.trim();
  }
  return ensureString(item.Name).trim() || 'Jellyfin Item';
}

function shouldPreferDirectPlay(source: JellyfinMediaSource, config: JellyfinConfig): boolean {
  if (source.SupportsDirectStream !== true) return false;
  if (config.directPlayPreferred === false) return false;

  const container = ensureString(source.Container).toLowerCase();
  const allowlist = Array.isArray(config.directPlayContainers)
    ? config.directPlayContainers.map((entry) => entry.toLowerCase())
    : [];
  if (!container || allowlist.length === 0) return true;
  return allowlist.includes(container);
}

export async function authenticateWithPassword(
  serverUrl: string,
  username: string,
  password: string,
  client: JellyfinClientInfo,
): Promise<JellyfinAuthSession> {
  const normalizedUrl = normalizeBaseUrl(serverUrl);
  if (!normalizedUrl) throw new Error('Missing Jellyfin server URL.');
  if (!username.trim()) throw new Error('Missing Jellyfin username.');
  if (!password) throw new Error('Missing Jellyfin password.');

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), JELLYFIN_LOGIN_TIMEOUT_MS);
  let response: Response;
  try {
    response = await fetch(`${normalizedUrl}/Users/AuthenticateByName`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: createAuthorizationHeader(client),
      },
      body: JSON.stringify({
        Username: username,
        Pw: password,
      }),
      signal: controller.signal,
    });
  } catch (error) {
    if (isAbortError(error)) {
      throw new Error('Jellyfin login timed out. Check the server URL and network connection.');
    }
    throw new Error(`Could not reach Jellyfin server (${getErrorMessage(error)}).`);
  } finally {
    clearTimeout(timeout);
  }

  if (response.status === 401 || response.status === 403) {
    throw new Error('Invalid Jellyfin username or password.');
  }
  if (!response.ok) {
    throw new Error(`Jellyfin login failed (${response.status} ${response.statusText}).`);
  }

  const payload = (await response.json()) as JellyfinAuthResponse;
  const accessToken = ensureString(payload.AccessToken);
  const userId = ensureString(payload.User?.Id);
  if (!accessToken || !userId) {
    throw new Error('Jellyfin login response missing token/user.');
  }

  return {
    serverUrl: normalizedUrl,
    accessToken,
    userId,
    username: username.trim(),
  };
}

export async function listLibraries(
  session: JellyfinAuthSession,
  client: JellyfinClientInfo,
): Promise<JellyfinLibrary[]> {
  const payload = await jellyfinRequestJson<JellyfinItemsResponse>(
    `/Users/${session.userId}/Views`,
    { method: 'GET' },
    session,
    client,
  );

  const items = Array.isArray(payload.Items) ? payload.Items : [];
  return items.map((item) => ({
    id: ensureString(item.Id),
    name: ensureString(item.Name, 'Untitled'),
    collectionType: ensureString((item as { CollectionType?: string }).CollectionType),
    type: ensureString(item.Type),
  }));
}

export async function listItems(
  session: JellyfinAuthSession,
  client: JellyfinClientInfo,
  options: {
    libraryId: string;
    searchTerm?: string;
    limit?: number;
    recursive?: boolean;
    includeItemTypes?: string;
  },
): Promise<Array<{ id: string; name: string; type: string; title: string }>> {
  if (!options.libraryId) throw new Error('Missing Jellyfin library id.');
  const normalizedSearchTerm = options.searchTerm?.trim() || '';
  const includeItemTypes =
    options.includeItemTypes?.trim() ||
    (normalizedSearchTerm
      ? 'Movie,Episode,Audio,Series,Season,Folder,CollectionFolder'
      : 'Movie,Episode,Audio');

  const query = new URLSearchParams({
    ParentId: options.libraryId,
    Recursive: options.recursive === false ? 'false' : 'true',
    IncludeItemTypes: includeItemTypes,
    Fields: 'MediaSources,UserData',
    SortBy: 'SortName',
    SortOrder: 'Ascending',
    Limit: String(options.limit ?? 100),
  });
  if (normalizedSearchTerm) {
    query.set('SearchTerm', normalizedSearchTerm);
  }

  const payload = await jellyfinRequestJson<JellyfinItemsResponse>(
    `/Users/${session.userId}/Items?${query.toString()}`,
    { method: 'GET' },
    session,
    client,
  );
  const items = Array.isArray(payload.Items) ? payload.Items : [];
  return items.map((item) => ({
    id: ensureString(item.Id),
    name: ensureString(item.Name),
    type: ensureString(item.Type),
    title: getDisplayTitle(item),
  }));
}

export async function listSubtitleTracks(
  session: JellyfinAuthSession,
  client: JellyfinClientInfo,
  itemId: string,
): Promise<JellyfinSubtitleTrack[]> {
  if (!itemId.trim()) throw new Error('Missing Jellyfin item id.');
  let source: JellyfinMediaSource | undefined;

  try {
    const playbackInfo = await jellyfinRequestJson<JellyfinPlaybackInfoResponse>(
      `/Items/${itemId}/PlaybackInfo?UserId=${encodeURIComponent(session.userId)}`,
      {
        method: 'POST',
        body: JSON.stringify({ UserId: session.userId }),
      },
      session,
      client,
    );
    source = Array.isArray(playbackInfo.MediaSources) ? playbackInfo.MediaSources[0] : undefined;
  } catch {}

  if (!source) {
    const item = await jellyfinRequestJson<JellyfinItem>(
      `/Users/${session.userId}/Items/${itemId}?Fields=MediaSources`,
      { method: 'GET' },
      session,
      client,
    );
    source = Array.isArray(item.MediaSources) ? item.MediaSources[0] : undefined;
  }

  if (!source) {
    throw new Error('No playable media source found for Jellyfin item.');
  }
  const mediaSourceId = ensureString(source.Id);

  const streams = Array.isArray(source.MediaStreams) ? source.MediaStreams : [];
  const tracks: JellyfinSubtitleTrack[] = [];
  for (const stream of streams) {
    if (stream.Type !== 'Subtitle') continue;
    const index = asIntegerOrNull(stream.Index);
    if (index === null) continue;
    tracks.push({
      index,
      language: ensureString(stream.Language),
      title: ensureString(stream.DisplayTitle || stream.Title),
      codec: ensureString(stream.Codec),
      isDefault: stream.IsDefault === true,
      isForced: stream.IsForced === true,
      isExternal: stream.IsExternal === true,
      deliveryMethod: ensureString(stream.DeliveryMethod),
      deliveryUrl: resolveDeliveryUrl(session, stream, itemId, mediaSourceId),
    });
  }
  return tracks;
}

export async function resolvePlaybackPlan(
  session: JellyfinAuthSession,
  client: JellyfinClientInfo,
  config: JellyfinConfig,
  selection: JellyfinPlaybackSelection,
): Promise<JellyfinPlaybackPlan> {
  if (!selection.itemId) {
    throw new Error('Missing Jellyfin item id.');
  }

  const item = await jellyfinRequestJson<JellyfinItem>(
    `/Users/${session.userId}/Items/${selection.itemId}?Fields=MediaSources,UserData`,
    { method: 'GET' },
    session,
    client,
  );
  const source = Array.isArray(item.MediaSources) ? item.MediaSources[0] : undefined;
  if (!source) {
    throw new Error('No playable media source found for Jellyfin item.');
  }

  const defaults = getStreamDefaults(source);
  const audioStreamIndex = selection.audioStreamIndex ?? defaults.audioStreamIndex ?? null;
  const subtitleStreamIndex = selection.subtitleStreamIndex ?? null;
  const startTimeTicks = Math.max(0, asIntegerOrNull(item.UserData?.PlaybackPositionTicks) ?? 0);
  const basePlan: JellyfinPlaybackPlan = {
    mode: 'transcode',
    url: '',
    title: getDisplayTitle(item),
    startTimeTicks,
    audioStreamIndex,
    subtitleStreamIndex,
  };

  if (shouldPreferDirectPlay(source, config)) {
    return {
      ...basePlan,
      mode: 'direct',
      url: createDirectPlayUrl(session, selection.itemId, source, basePlan),
    };
  }
  if (source.SupportsTranscoding !== true && source.SupportsDirectStream === true) {
    return {
      ...basePlan,
      mode: 'direct',
      url: createDirectPlayUrl(session, selection.itemId, source, basePlan),
    };
  }
  if (source.SupportsTranscoding !== true) {
    throw new Error('Jellyfin item cannot be streamed by direct play or transcoding.');
  }

  return {
    ...basePlan,
    mode: 'transcode',
    url: createTranscodeUrl(session, selection.itemId, source, basePlan, config),
  };
}

export function ticksToSeconds(ticks: number): number {
  return Math.max(0, Math.floor(ticks / JELLYFIN_TICKS_PER_SECOND));
}
