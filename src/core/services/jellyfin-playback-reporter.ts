export interface JellyfinTimelinePlaybackState {
  itemId: string;
  mediaSourceId?: string;
  positionTicks?: number;
  playbackStartTimeTicks?: number;
  isPaused?: boolean;
  isMuted?: boolean;
  canSeek?: boolean;
  volumeLevel?: number;
  playbackRate?: number;
  playMethod?: string;
  audioStreamIndex?: number | null;
  subtitleStreamIndex?: number | null;
  playlistItemId?: string | null;
  eventName?: string;
  failed?: boolean;
}

export interface JellyfinTimelinePayload {
  ItemId: string;
  MediaSourceId?: string;
  PositionTicks: number;
  PlaybackStartTimeTicks: number;
  IsPaused: boolean;
  IsMuted: boolean;
  CanSeek: boolean;
  VolumeLevel: number;
  PlaybackRate: number;
  PlayMethod: string;
  AudioStreamIndex?: number | null;
  SubtitleStreamIndex?: number | null;
  PlaylistItemId?: string | null;
  Failed?: boolean;
}

export interface JellyfinPlaybackReporterOptions {
  serverUrl: string;
  accessToken: string;
  deviceId: string;
  clientName?: string;
  clientVersion?: string;
  deviceName?: string;
  fetchImpl?: typeof fetch;
  logWarn?: (message: string, details?: unknown) => void;
}

function clampVolume(value: number | undefined): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return 100;
  return Math.max(0, Math.min(100, Math.round(value)));
}

function normalizeTicks(value: number | undefined): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return 0;
  return Math.max(0, Math.floor(value));
}

function asNullableInteger(value: number | null | undefined): number | null {
  if (typeof value !== 'number' || !Number.isInteger(value)) return null;
  return value;
}

export function buildJellyfinTimelinePayload(
  state: JellyfinTimelinePlaybackState,
): JellyfinTimelinePayload {
  return {
    ItemId: state.itemId,
    MediaSourceId: state.mediaSourceId,
    PositionTicks: normalizeTicks(state.positionTicks),
    PlaybackStartTimeTicks: normalizeTicks(state.playbackStartTimeTicks),
    IsPaused: state.isPaused === true,
    IsMuted: state.isMuted === true,
    CanSeek: state.canSeek !== false,
    VolumeLevel: clampVolume(state.volumeLevel),
    PlaybackRate:
      typeof state.playbackRate === 'number' && Number.isFinite(state.playbackRate)
        ? state.playbackRate
        : 1,
    PlayMethod: state.playMethod || 'DirectPlay',
    AudioStreamIndex: asNullableInteger(state.audioStreamIndex),
    SubtitleStreamIndex: asNullableInteger(state.subtitleStreamIndex),
    PlaylistItemId: state.playlistItemId,
    Failed: state.failed,
  };
}

// Authenticated JSON POSTs as this device. Playback timeline reports (start, progress, stop)
// are plain HTTP, so they work whether or not the cast-control websocket is connected; that
// is how Jellyfin learns the resume position and marks items played for every playback.
export class JellyfinPlaybackReporter {
  readonly authHeader: string;
  private readonly serverUrl: string;
  private readonly fetchImpl: typeof fetch;
  private readonly logWarn?: (message: string, details?: unknown) => void;
  private readonly failedRequestPaths = new Set<string>();

  constructor(options: JellyfinPlaybackReporterOptions) {
    this.serverUrl = options.serverUrl.trim().replace(/\/+$/, '');
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.logWarn = options.logWarn;
    const clientName = options.clientName || 'SubMiner';
    const clientVersion = options.clientVersion || '0.1.0';
    const deviceName = options.deviceName || clientName;
    this.authHeader = `MediaBrowser Client="${clientName}", Device="${deviceName}", DeviceId="${options.deviceId}", Version="${clientVersion}", Token="${options.accessToken}"`;
  }

  reportPlaying(state: JellyfinTimelinePlaybackState): Promise<boolean> {
    return this.postJson('/Sessions/Playing', buildJellyfinTimelinePayload(state));
  }

  reportProgress(state: JellyfinTimelinePlaybackState): Promise<boolean> {
    return this.postJson('/Sessions/Playing/Progress', buildJellyfinTimelinePayload(state));
  }

  reportStopped(state: JellyfinTimelinePlaybackState): Promise<boolean> {
    return this.postJson('/Sessions/Playing/Stopped', {
      ...buildJellyfinTimelinePayload(state),
      Failed: state.failed === true,
    });
  }

  async postJson(path: string, payload: unknown): Promise<boolean> {
    try {
      const response = await this.fetchImpl(`${this.serverUrl}${path}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: this.authHeader,
        },
        body: JSON.stringify(payload),
      });
      this.noteRequestOutcome(path, response.ok ? null : `HTTP ${response.status}`);
      return response.ok;
    } catch (error) {
      this.noteRequestOutcome(path, error);
      return false;
    }
  }

  // Warn once per path while it keeps failing so a rejected stop report is visible in the
  // log without a warning per progress tick.
  private noteRequestOutcome(path: string, failure: unknown): void {
    if (failure === null) {
      this.failedRequestPaths.delete(path);
      return;
    }
    if (this.failedRequestPaths.has(path)) return;
    this.failedRequestPaths.add(path);
    this.logWarn?.(`Jellyfin request failed: POST ${path}`, failure);
  }
}
