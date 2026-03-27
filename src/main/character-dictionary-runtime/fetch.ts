import type { AnilistMediaGuess } from '../../core/services/anilist/anilist-updater';
import { ANILIST_GRAPHQL_URL } from './constants';
import type {
  CharacterDictionaryRole,
  CharacterRecord,
  ResolvedAniListMedia,
  VoiceActorRecord,
} from './types';

type AniListSearchResponse = {
  Page?: {
    media?: Array<{
      id: number;
      episodes?: number | null;
      title?: {
        romaji?: string | null;
        english?: string | null;
        native?: string | null;
      };
    }>;
  };
};

type AniListCharacterPageResponse = {
  Media?: {
    title?: {
      romaji?: string | null;
      english?: string | null;
      native?: string | null;
    };
    characters?: {
      pageInfo?: {
        hasNextPage?: boolean | null;
      };
      edges?: Array<{
        role?: string | null;
        voiceActors?: Array<{
          id: number;
          name?: {
            full?: string | null;
            native?: string | null;
          } | null;
          image?: {
            large?: string | null;
            medium?: string | null;
          } | null;
        }> | null;
        node?: {
          id: number;
          description?: string | null;
          image?: {
            large?: string | null;
            medium?: string | null;
          } | null;
          gender?: string | null;
          age?: string | number | null;
          dateOfBirth?: {
            month?: number | null;
            day?: number | null;
          } | null;
          bloodType?: string | null;
          name?: {
            first?: string | null;
            full?: string | null;
            last?: string | null;
            native?: string | null;
            alternative?: Array<string | null> | null;
          } | null;
        } | null;
      } | null>;
    } | null;
  } | null;
};

function normalizeTitle(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, ' ');
}

function pickAniListSearchResult(
  title: string,
  episode: number | null,
  media: Array<{
    id: number;
    episodes?: number | null;
    title?: {
      romaji?: string | null;
      english?: string | null;
      native?: string | null;
    };
  }>,
): ResolvedAniListMedia | null {
  if (media.length === 0) return null;

  const episodeFiltered =
    episode && episode > 0
      ? media.filter((entry) => {
          const totalEpisodes = entry.episodes;
          return (
            typeof totalEpisodes !== 'number' || totalEpisodes <= 0 || episode <= totalEpisodes
          );
        })
      : media;
  const candidates = episodeFiltered.length > 0 ? episodeFiltered : media;
  const normalizedTitle = normalizeTitle(title);

  const exact = candidates.find((entry) => {
    const titles = [entry.title?.english, entry.title?.romaji, entry.title?.native]
      .filter((value): value is string => typeof value === 'string')
      .map((value) => normalizeTitle(value));
    return titles.includes(normalizedTitle);
  });
  const selected = exact ?? candidates[0] ?? media[0];
  if (!selected) return null;

  const selectedTitle =
    selected.title?.english?.trim() ||
    selected.title?.romaji?.trim() ||
    selected.title?.native?.trim() ||
    title.trim();
  return {
    id: selected.id,
    title: selectedTitle,
  };
}

async function fetchAniList<T>(
  query: string,
  variables: Record<string, unknown>,
  beforeRequest?: () => Promise<void>,
): Promise<T> {
  if (beforeRequest) {
    await beforeRequest();
  }
  const response = await fetch(ANILIST_GRAPHQL_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      query,
      variables,
    }),
  });
  if (!response.ok) {
    throw new Error(`AniList request failed (${response.status})`);
  }
  const payload = (await response.json()) as {
    data?: T;
    errors?: Array<{ message?: string }>;
  };
  const firstError = payload.errors?.find((entry) => entry && typeof entry.message === 'string');
  if (firstError?.message) {
    throw new Error(firstError.message);
  }
  if (!payload.data) {
    throw new Error('AniList response missing data');
  }
  return payload.data;
}

function mapRole(input: string | null | undefined): CharacterDictionaryRole {
  const value = (input || '').trim().toUpperCase();
  if (value === 'MAIN') return 'main';
  if (value === 'SUPPORTING') return 'primary';
  if (value === 'BACKGROUND') return 'side';
  return 'side';
}

function inferImageExt(contentType: string | null): string {
  const normalized = (contentType || '').toLowerCase();
  if (normalized.includes('png')) return 'png';
  if (normalized.includes('gif')) return 'gif';
  if (normalized.includes('webp')) return 'webp';
  return 'jpg';
}

export async function resolveAniListMediaIdFromGuess(
  guess: AnilistMediaGuess,
  beforeRequest?: () => Promise<void>,
): Promise<ResolvedAniListMedia> {
  const data = await fetchAniList<AniListSearchResponse>(
    `
      query($search: String!) {
        Page(perPage: 10) {
          media(search: $search, type: ANIME, sort: [SEARCH_MATCH, POPULARITY_DESC]) {
            id
            episodes
            title {
              romaji
              english
              native
            }
          }
        }
      }
    `,
    {
      search: guess.title,
    },
    beforeRequest,
  );

  const media = data.Page?.media ?? [];
  const resolved = pickAniListSearchResult(guess.title, guess.episode, media);
  if (!resolved) {
    throw new Error(`No AniList media match found for "${guess.title}".`);
  }
  return resolved;
}

export async function fetchCharactersForMedia(
  mediaId: number,
  beforeRequest?: () => Promise<void>,
  onPageFetched?: (page: number) => void,
): Promise<{
  mediaTitle: string;
  characters: CharacterRecord[];
}> {
  const characters: CharacterRecord[] = [];
  let page = 1;
  let mediaTitle = '';
  for (;;) {
    const data = await fetchAniList<AniListCharacterPageResponse>(
      `
        query($id: Int!, $page: Int!) {
          Media(id: $id, type: ANIME) {
            title {
              romaji
              english
              native
            }
            characters(page: $page, perPage: 50, sort: [ROLE, RELEVANCE, ID]) {
              pageInfo {
                hasNextPage
              }
              edges {
                role
                voiceActors(language: JAPANESE) {
                  id
                  name {
                    full
                    native
                  }
                  image {
                    medium
                  }
                }
                node {
                  id
                  description(asHtml: false)
                  gender
                  age
                  dateOfBirth {
                    month
                    day
                  }
                  bloodType
                  image {
                    large
                    medium
                  }
                  name {
                    first
                    full
                    last
                    native
                    alternative
                  }
                }
              }
            }
          }
        }
      `,
      {
        id: mediaId,
        page,
      },
      beforeRequest,
    );
    onPageFetched?.(page);

    const media = data.Media;
    if (!media) {
      throw new Error(`AniList media ${mediaId} not found.`);
    }
    if (!mediaTitle) {
      mediaTitle =
        media.title?.english?.trim() ||
        media.title?.romaji?.trim() ||
        media.title?.native?.trim() ||
        `AniList ${mediaId}`;
    }

    const edges = media.characters?.edges ?? [];
    for (const edge of edges) {
      const node = edge?.node;
      if (!node || typeof node.id !== 'number') continue;
      const firstNameHint = node.name?.first?.trim() || '';
      const fullName = node.name?.full?.trim() || '';
      const lastNameHint = node.name?.last?.trim() || '';
      const nativeName = node.name?.native?.trim() || '';
      const alternativeNames = [
        ...new Set(
          (node.name?.alternative ?? [])
            .filter((value): value is string => typeof value === 'string')
            .map((value) => value.trim())
            .filter((value) => value.length > 0),
        ),
      ];
      if (!nativeName) continue;
      const voiceActors: VoiceActorRecord[] = [];
      for (const va of edge?.voiceActors ?? []) {
        if (!va || typeof va.id !== 'number') continue;
        const vaFull = va.name?.full?.trim() || '';
        const vaNative = va.name?.native?.trim() || '';
        if (!vaFull && !vaNative) continue;
        voiceActors.push({
          id: va.id,
          fullName: vaFull,
          nativeName: vaNative,
          imageUrl: va.image?.medium || null,
        });
      }
      characters.push({
        id: node.id,
        role: mapRole(edge?.role),
        firstNameHint,
        fullName,
        lastNameHint,
        nativeName,
        alternativeNames,
        bloodType: node.bloodType?.trim() || '',
        birthday:
          typeof node.dateOfBirth?.month === 'number' && typeof node.dateOfBirth?.day === 'number'
            ? [node.dateOfBirth.month, node.dateOfBirth.day]
            : null,
        description: node.description || '',
        imageUrl: node.image?.large || node.image?.medium || null,
        age:
          typeof node.age === 'string'
            ? node.age.trim()
            : typeof node.age === 'number'
              ? String(node.age)
              : '',
        sex: node.gender?.trim() || '',
        voiceActors,
      });
    }

    const hasNextPage = Boolean(media.characters?.pageInfo?.hasNextPage);
    if (!hasNextPage) {
      break;
    }
    page += 1;
  }

  return {
    mediaTitle,
    characters,
  };
}

export async function downloadCharacterImage(
  imageUrl: string,
  charId: number,
): Promise<{
  filename: string;
  ext: string;
  bytes: Buffer;
} | null> {
  try {
    const response = await fetch(imageUrl);
    if (!response.ok) return null;
    const bytes = Buffer.from(await response.arrayBuffer());
    if (bytes.length === 0) return null;
    const ext = inferImageExt(response.headers.get('content-type'));
    return {
      filename: `c${charId}.${ext}`,
      ext,
      bytes,
    };
  } catch {
    return null;
  }
}
