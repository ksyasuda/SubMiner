/**
 * Wire types for the M-Extension-Server bridge, which runs Aniyomi
 * (`eu.kanade.tachiyomi.animeextension`) APKs on a desktop JVM and exposes
 * them over loopback HTTP.
 *
 * Field names mirror the server's JSON exactly, including Kotlin/OkHttp
 * internals that leak into the payload (see `OkHttpHeaders`).
 */

/** Marker key the server uses to select a source inside a SourceFactory APK. */
export const BRIDGE_CONTEXT_KEY = '__mangatan_bridge_context__';

/** Handshake shape from `GET /capabilities`. */
export interface BridgeCapabilities {
  mangatanMihonBridge?: number;
  sourceFactory?: boolean;
  preferenceCallbacks?: boolean;
  youtubeResolver?: boolean;
}

/**
 * OkHttp serializes `Headers` as a flat alternating name/value array under an
 * internal field name. Kept verbatim so parsing stays honest about the source.
 */
export interface OkHttpHeaders {
  namesAndValues$okhttp?: string[];
}

export interface BridgeTrack {
  url?: string;
  lang?: string;
}

/** One playable stream returned by an extension's `getVideoList`. */
export interface BridgeVideo {
  /** Page/embed URL the stream was extracted from. */
  url?: string;
  /** Display label, e.g. "1080p". */
  quality?: string;
  /**
   * Playable media URL. Normally a `/video/<token>` proxy URL on the bridge
   * itself, valid only while that server process is alive.
   */
  videoUrl?: string;
  headers?: OkHttpHeaders;
  audioTracks?: BridgeTrack[];
  subtitleTracks?: BridgeTrack[];
}

export interface BridgeEpisode {
  name?: string;
  url?: string;
  date_upload?: number;
  scanlator?: string;
  episode_number?: number;
}

export interface BridgeAnime {
  url?: string;
  title?: string;
  artist?: string;
  author?: string;
  description?: string;
  genres?: string[];
  status?: number;
  thumbnail_url?: string;
}

/** One source inside an extension APK. A SourceFactory APK yields several. */
export interface BridgeSourceDescriptor {
  id?: string | number;
  name?: string;
  lang?: string;
  baseUrl?: string;
}

export interface BridgeAnimePage {
  animes?: BridgeAnime[];
  hasNextPage?: boolean;
}

/** The server reports failures as HTTP 200 with an error body. */
export interface BridgeErrorBody {
  error?: string;
  code?: number;
}

/** A source preference entry, passed through to the extension unchanged. */
export interface BridgePreference {
  key: string;
  [field: string]: unknown;
}

/** Normalized stream, ready to hand to mpv. */
export interface ResolvedStream {
  url: string;
  quality: string;
  headers: Record<string, string>;
  subtitles: Array<{ url: string; lang: string }>;
  audios: Array<{ url: string; lang: string }>;
}
