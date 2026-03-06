import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createHash } from 'node:crypto';
import type { AnilistMediaGuess } from '../core/services/anilist/anilist-updater';
import { hasVideoExtension } from '../shared/video-extensions';

const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';
const ANILIST_REQUEST_DELAY_MS = 2000;
const CHARACTER_IMAGE_DOWNLOAD_DELAY_MS = 250;
const HONORIFIC_SUFFIXES = [
  'さん',
  '様',
  '先生',
  '先輩',
  '後輩',
  '氏',
  '君',
  'くん',
  'ちゃん',
  'たん',
  '坊',
  '殿',
  '博士',
  '社長',
  '部長',
] as const;
type CharacterDictionaryRole = 'main' | 'primary' | 'side' | 'appears';

type CharacterDictionaryGlossaryEntry = string | Record<string, unknown>;
type CharacterDictionaryTermEntry = [
  string,
  string,
  string,
  string,
  number,
  CharacterDictionaryGlossaryEntry[],
  number,
  string,
];

type CharacterDictionarySnapshotImage = {
  path: string;
  dataBase64: string;
};

export type CharacterDictionarySnapshot = {
  formatVersion: number;
  mediaId: number;
  mediaTitle: string;
  entryCount: number;
  updatedAt: number;
  termEntries: CharacterDictionaryTermEntry[];
  images: CharacterDictionarySnapshotImage[];
};

const CHARACTER_DICTIONARY_FORMAT_VERSION = 10;
const CHARACTER_DICTIONARY_MERGED_TITLE = 'SubMiner Character Dictionary';

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
        node?: {
          id: number;
          description?: string | null;
          image?: {
            large?: string | null;
            medium?: string | null;
          } | null;
          name?: {
            full?: string | null;
            native?: string | null;
          } | null;
        } | null;
      } | null>;
    } | null;
  } | null;
};

type CharacterRecord = {
  id: number;
  role: CharacterDictionaryRole;
  fullName: string;
  nativeName: string;
  description: string;
  imageUrl: string | null;
};

type ZipEntry = {
  name: string;
  data: Buffer;
  crc32: number;
  localHeaderOffset: number;
};

export type CharacterDictionaryBuildResult = {
  zipPath: string;
  fromCache: boolean;
  mediaId: number;
  mediaTitle: string;
  entryCount: number;
  dictionaryTitle?: string;
  revision?: string;
};

export type CharacterDictionaryGenerateOptions = {
  refreshTtlMs?: number;
};

export type CharacterDictionarySnapshotResult = {
  mediaId: number;
  mediaTitle: string;
  entryCount: number;
  fromCache: boolean;
  updatedAt: number;
};

export type MergedCharacterDictionaryBuildResult = {
  zipPath: string;
  revision: string;
  dictionaryTitle: string;
  entryCount: number;
};

export interface CharacterDictionaryRuntimeDeps {
  userDataPath: string;
  getCurrentMediaPath: () => string | null;
  getCurrentMediaTitle: () => string | null;
  resolveMediaPathForJimaku: (mediaPath: string | null) => string | null;
  guessAnilistMediaInfo: (
    mediaPath: string | null,
    mediaTitle: string | null,
  ) => Promise<AnilistMediaGuess | null>;
  now: () => number;
  sleep?: (ms: number) => Promise<void>;
  logInfo?: (message: string) => void;
  logWarn?: (message: string) => void;
}

type ResolvedAniListMedia = {
  id: number;
  title: string;
};

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

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
    typeof episode === 'number' && episode > 0
      ? media.filter((entry) => entry.episodes == null || entry.episodes >= episode)
      : media;
  const candidates = episodeFiltered.length > 0 ? episodeFiltered : media;
  const normalizedInput = normalizeTitle(title);
  const exact = candidates.find((entry) => {
    const candidateTitles = [entry.title?.romaji, entry.title?.english, entry.title?.native]
      .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
      .map((value) => normalizeTitle(value));
    return candidateTitles.includes(normalizedInput);
  });
  const selected = exact ?? candidates[0]!;
  const selectedTitle =
    selected.title?.english?.trim() ||
    selected.title?.romaji?.trim() ||
    selected.title?.native?.trim() ||
    title;
  return {
    id: selected.id,
    title: selectedTitle,
  };
}

function hasKanaOnly(value: string): boolean {
  return /^[\u3040-\u309f\u30a0-\u30ffー]+$/.test(value);
}

function katakanaToHiragana(value: string): string {
  let output = '';
  for (const char of value) {
    const code = char.charCodeAt(0);
    if (code >= 0x30a1 && code <= 0x30f6) {
      output += String.fromCharCode(code - 0x60);
      continue;
    }
    output += char;
  }
  return output;
}

function buildReading(term: string): string {
  const compact = term.replace(/\s+/g, '').trim();
  if (!compact || !hasKanaOnly(compact)) {
    return '';
  }
  return katakanaToHiragana(compact);
}

function isRomanizedName(value: string): boolean {
  return /^[A-Za-zĀĪŪĒŌÂÊÎÔÛāīūēōâêîôû'’.\-\s]+$/.test(value);
}

function normalizeRomanizedName(value: string): string {
  return value
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[’']/g, '')
    .replace(/[.\-]/g, ' ')
    .replace(/ā|â/g, 'aa')
    .replace(/ī|î/g, 'ii')
    .replace(/ū|û/g, 'uu')
    .replace(/ē|ê/g, 'ei')
    .replace(/ō|ô/g, 'ou')
    .replace(/\s+/g, ' ')
    .trim();
}

const ROMANIZED_KANA_DIGRAPHS: ReadonlyArray<[string, string]> = [
  ['kya', 'キャ'],
  ['kyu', 'キュ'],
  ['kyo', 'キョ'],
  ['gya', 'ギャ'],
  ['gyu', 'ギュ'],
  ['gyo', 'ギョ'],
  ['sha', 'シャ'],
  ['shu', 'シュ'],
  ['sho', 'ショ'],
  ['sya', 'シャ'],
  ['syu', 'シュ'],
  ['syo', 'ショ'],
  ['ja', 'ジャ'],
  ['ju', 'ジュ'],
  ['jo', 'ジョ'],
  ['jya', 'ジャ'],
  ['jyu', 'ジュ'],
  ['jyo', 'ジョ'],
  ['cha', 'チャ'],
  ['chu', 'チュ'],
  ['cho', 'チョ'],
  ['tya', 'チャ'],
  ['tyu', 'チュ'],
  ['tyo', 'チョ'],
  ['cya', 'チャ'],
  ['cyu', 'チュ'],
  ['cyo', 'チョ'],
  ['nya', 'ニャ'],
  ['nyu', 'ニュ'],
  ['nyo', 'ニョ'],
  ['hya', 'ヒャ'],
  ['hyu', 'ヒュ'],
  ['hyo', 'ヒョ'],
  ['bya', 'ビャ'],
  ['byu', 'ビュ'],
  ['byo', 'ビョ'],
  ['pya', 'ピャ'],
  ['pyu', 'ピュ'],
  ['pyo', 'ピョ'],
  ['mya', 'ミャ'],
  ['myu', 'ミュ'],
  ['myo', 'ミョ'],
  ['rya', 'リャ'],
  ['ryu', 'リュ'],
  ['ryo', 'リョ'],
  ['fa', 'ファ'],
  ['fi', 'フィ'],
  ['fe', 'フェ'],
  ['fo', 'フォ'],
  ['fyu', 'フュ'],
  ['fyo', 'フョ'],
  ['fya', 'フャ'],
  ['va', 'ヴァ'],
  ['vi', 'ヴィ'],
  ['vu', 'ヴ'],
  ['ve', 'ヴェ'],
  ['vo', 'ヴォ'],
  ['she', 'シェ'],
  ['che', 'チェ'],
  ['je', 'ジェ'],
  ['tsi', 'ツィ'],
  ['tse', 'ツェ'],
  ['tsa', 'ツァ'],
  ['tso', 'ツォ'],
  ['thi', 'ティ'],
  ['thu', 'テュ'],
  ['dhi', 'ディ'],
  ['dhu', 'デュ'],
  ['wi', 'ウィ'],
  ['we', 'ウェ'],
  ['wo', 'ウォ'],
];

const ROMANIZED_KANA_MONOGRAPHS: ReadonlyArray<[string, string]> = [
  ['a', 'ア'],
  ['i', 'イ'],
  ['u', 'ウ'],
  ['e', 'エ'],
  ['o', 'オ'],
  ['ka', 'カ'],
  ['ki', 'キ'],
  ['ku', 'ク'],
  ['ke', 'ケ'],
  ['ko', 'コ'],
  ['ga', 'ガ'],
  ['gi', 'ギ'],
  ['gu', 'グ'],
  ['ge', 'ゲ'],
  ['go', 'ゴ'],
  ['sa', 'サ'],
  ['shi', 'シ'],
  ['si', 'シ'],
  ['su', 'ス'],
  ['se', 'セ'],
  ['so', 'ソ'],
  ['za', 'ザ'],
  ['ji', 'ジ'],
  ['zi', 'ジ'],
  ['zu', 'ズ'],
  ['ze', 'ゼ'],
  ['zo', 'ゾ'],
  ['ta', 'タ'],
  ['chi', 'チ'],
  ['ti', 'チ'],
  ['tsu', 'ツ'],
  ['tu', 'ツ'],
  ['te', 'テ'],
  ['to', 'ト'],
  ['da', 'ダ'],
  ['de', 'デ'],
  ['do', 'ド'],
  ['na', 'ナ'],
  ['ni', 'ニ'],
  ['nu', 'ヌ'],
  ['ne', 'ネ'],
  ['no', 'ノ'],
  ['ha', 'ハ'],
  ['hi', 'ヒ'],
  ['fu', 'フ'],
  ['hu', 'フ'],
  ['he', 'ヘ'],
  ['ho', 'ホ'],
  ['ba', 'バ'],
  ['bi', 'ビ'],
  ['bu', 'ブ'],
  ['be', 'ベ'],
  ['bo', 'ボ'],
  ['pa', 'パ'],
  ['pi', 'ピ'],
  ['pu', 'プ'],
  ['pe', 'ペ'],
  ['po', 'ポ'],
  ['ma', 'マ'],
  ['mi', 'ミ'],
  ['mu', 'ム'],
  ['me', 'メ'],
  ['mo', 'モ'],
  ['ya', 'ヤ'],
  ['yu', 'ユ'],
  ['yo', 'ヨ'],
  ['ra', 'ラ'],
  ['ri', 'リ'],
  ['ru', 'ル'],
  ['re', 'レ'],
  ['ro', 'ロ'],
  ['wa', 'ワ'],
  ['wo', 'ヲ'],
  ['n', 'ン'],
];

function romanizedTokenToKatakana(token: string): string | null {
  const normalized = normalizeRomanizedName(token).replace(/\s+/g, '');
  if (!normalized || !/^[a-z]+$/.test(normalized)) {
    return null;
  }

  let output = '';
  for (let i = 0; i < normalized.length; ) {
    const current = normalized[i]!;
    const next = normalized[i + 1] ?? '';

    if (
      i + 1 < normalized.length &&
      current === next &&
      current !== 'n' &&
      !'aeiou'.includes(current)
    ) {
      output += 'ッ';
      i += 1;
      continue;
    }

    if (
      current === 'n' &&
      next.length > 0 &&
      next !== 'y' &&
      !'aeiou'.includes(next)
    ) {
      output += 'ン';
      i += 1;
      continue;
    }

    const digraph = ROMANIZED_KANA_DIGRAPHS.find(([romaji]) =>
      normalized.startsWith(romaji, i),
    );
    if (digraph) {
      output += digraph[1];
      i += digraph[0].length;
      continue;
    }

    const monograph = ROMANIZED_KANA_MONOGRAPHS.find(([romaji]) =>
      normalized.startsWith(romaji, i),
    );
    if (monograph) {
      output += monograph[1];
      i += monograph[0].length;
      continue;
    }

    return null;
  }

  return output.length > 0 ? output : null;
}

function addRomanizedKanaAliases(values: Iterable<string>): string[] {
  const aliases = new Set<string>();
  for (const value of values) {
    const trimmed = value.trim();
    if (!trimmed || !isRomanizedName(trimmed)) continue;
    const katakana = romanizedTokenToKatakana(trimmed);
    if (katakana) {
      aliases.add(katakana);
    }
  }
  return [...aliases];
}

function buildNameTerms(character: CharacterRecord): string[] {
  const base = new Set<string>();
  const rawNames = [character.nativeName, character.fullName];
  for (const rawName of rawNames) {
    const name = rawName.trim();
    if (!name) continue;
    base.add(name);

    const compact = name.replace(/[\s\u3000]+/g, '');
    if (compact && compact !== name) {
      base.add(compact);
    }

    const noMiddleDots = compact.replace(/[・･·•]/g, '');
    if (noMiddleDots && noMiddleDots !== compact) {
      base.add(noMiddleDots);
    }

    const split = name.split(/[\s\u3000]+/).filter((part) => part.trim().length > 0);
    if (split.length === 2) {
      base.add(split[0]!);
      base.add(split[1]!);
    }

    const splitByMiddleDot = name
      .split(/[・･·•]/)
      .map((part) => part.trim())
      .filter((part) => part.length > 0);
    if (splitByMiddleDot.length >= 2) {
      for (const part of splitByMiddleDot) {
        base.add(part);
      }
    }
  }

  const withHonorifics = new Set<string>();
  for (const entry of base) {
    withHonorifics.add(entry);
    for (const suffix of HONORIFIC_SUFFIXES) {
      withHonorifics.add(`${entry}${suffix}`);
    }
  }

  for (const alias of addRomanizedKanaAliases(withHonorifics)) {
    withHonorifics.add(alias);
    for (const suffix of HONORIFIC_SUFFIXES) {
      withHonorifics.add(`${alias}${suffix}`);
    }
  }

  return [...withHonorifics].filter((entry) => entry.trim().length > 0);
}

function stripDescription(value: string): string {
  return value.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function normalizeDescription(value: string): string {
  const stripped = stripDescription(value);
  if (!stripped) return '';
  return stripped
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '$1')
    .replace(/https?:\/\/\S+/g, '')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/~!/g, '')
    .replace(/!~/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function roleInfo(role: CharacterDictionaryRole): { tag: string; score: number } {
  if (role === 'main') return { tag: 'main', score: 100 };
  if (role === 'primary') return { tag: 'primary', score: 75 };
  if (role === 'side') return { tag: 'side', score: 50 };
  return { tag: 'appears', score: 25 };
}

function mapRole(input: string | null | undefined): CharacterDictionaryRole {
  const value = (input || '').trim().toUpperCase();
  if (value === 'MAIN') return 'main';
  if (value === 'BACKGROUND') return 'appears';
  if (value === 'SUPPORTING') return 'side';
  return 'primary';
}

function roleLabel(role: CharacterDictionaryRole): string {
  if (role === 'main') return 'Main';
  if (role === 'primary') return 'Primary';
  if (role === 'side') return 'Side';
  return 'Appears';
}

function inferImageExt(contentType: string | null): string {
  const normalized = (contentType || '').toLowerCase();
  if (normalized.includes('png')) return 'png';
  if (normalized.includes('gif')) return 'gif';
  if (normalized.includes('webp')) return 'webp';
  return 'jpg';
}

function ensureDir(dirPath: string): void {
  if (fs.existsSync(dirPath)) return;
  fs.mkdirSync(dirPath, { recursive: true });
}

function expandUserPath(input: string): string {
  if (input.startsWith('~')) {
    return path.join(os.homedir(), input.slice(1));
  }
  return input;
}

function isVideoFile(filePath: string): boolean {
  return hasVideoExtension(path.extname(filePath));
}

function findFirstVideoFileInDirectory(directoryPath: string): string | null {
  const queue: string[] = [directoryPath];
  while (queue.length > 0) {
    const current = queue.shift()!;
    let entries: fs.Dirent[] = [];
    try {
      entries = fs.readdirSync(current, { withFileTypes: true });
    } catch {
      continue;
    }
    entries.sort((a, b) => a.name.localeCompare(b.name));
    for (const entry of entries) {
      const fullPath = path.join(current, entry.name);
      if (entry.isFile() && isVideoFile(fullPath)) {
        return fullPath;
      }
      if (entry.isDirectory() && !entry.name.startsWith('.')) {
        queue.push(fullPath);
      }
    }
  }
  return null;
}

function resolveDictionaryGuessInputs(targetPath: string): {
  mediaPath: string;
  mediaTitle: string | null;
} {
  const trimmed = targetPath.trim();
  if (!trimmed) {
    throw new Error('Dictionary target path is empty.');
  }
  const resolvedPath = path.resolve(expandUserPath(trimmed));
  let stats: fs.Stats;
  try {
    stats = fs.statSync(resolvedPath);
  } catch {
    throw new Error(`Dictionary target path not found: ${targetPath}`);
  }

  if (stats.isFile()) {
    return {
      mediaPath: resolvedPath,
      mediaTitle: path.basename(resolvedPath),
    };
  }

  if (stats.isDirectory()) {
    const firstVideo = findFirstVideoFileInDirectory(resolvedPath);
    if (firstVideo) {
      return {
        mediaPath: firstVideo,
        mediaTitle: path.basename(firstVideo),
      };
    }
    return {
      mediaPath: resolvedPath,
      mediaTitle: path.basename(resolvedPath),
    };
  }

  throw new Error(`Dictionary target must be a file or directory path: ${targetPath}`);
}

function getSnapshotsDir(outputDir: string): string {
  return path.join(outputDir, 'snapshots');
}

function getSnapshotPath(outputDir: string, mediaId: number): string {
  return path.join(getSnapshotsDir(outputDir), `anilist-${mediaId}.json`);
}

function getMergedZipPath(outputDir: string): string {
  return path.join(outputDir, 'merged.zip');
}

function readSnapshot(snapshotPath: string): CharacterDictionarySnapshot | null {
  try {
    const raw = fs.readFileSync(snapshotPath, 'utf8');
    const parsed = JSON.parse(raw) as Partial<CharacterDictionarySnapshot>;
    if (!parsed || typeof parsed !== 'object') {
      return null;
    }
    if (
      parsed.formatVersion !== CHARACTER_DICTIONARY_FORMAT_VERSION ||
      typeof parsed.mediaId !== 'number' ||
      typeof parsed.mediaTitle !== 'string' ||
      typeof parsed.entryCount !== 'number' ||
      typeof parsed.updatedAt !== 'number' ||
      !Array.isArray(parsed.termEntries) ||
      !Array.isArray(parsed.images)
    ) {
      return null;
    }
    return {
      formatVersion: parsed.formatVersion,
      mediaId: parsed.mediaId,
      mediaTitle: parsed.mediaTitle,
      entryCount: parsed.entryCount,
      updatedAt: parsed.updatedAt,
      termEntries: parsed.termEntries as CharacterDictionaryTermEntry[],
      images: parsed.images as CharacterDictionarySnapshotImage[],
    };
  } catch {
    return null;
  }
}

function writeSnapshot(snapshotPath: string, snapshot: CharacterDictionarySnapshot): void {
  ensureDir(path.dirname(snapshotPath));
  fs.writeFileSync(snapshotPath, JSON.stringify(snapshot, null, 2), 'utf8');
}

function createDefinitionGlossary(
  character: CharacterRecord,
  mediaTitle: string,
  imagePath: string | null,
): CharacterDictionaryGlossaryEntry[] {
  const displayName = character.nativeName || character.fullName || `Character ${character.id}`;
  const lines: string[] = [`${displayName} [${roleLabel(character.role)}]`, `${mediaTitle} · AniList`];

  const description = normalizeDescription(character.description);
  if (description) {
    lines.push(description);
  }

  if (!imagePath) {
    return [lines.join('\n')];
  }

  const content: Array<string | Record<string, unknown>> = [
    {
      tag: 'img',
      path: imagePath,
      width: 8,
      height: 11,
      sizeUnits: 'em',
      title: displayName,
      alt: displayName,
      description: `${displayName} · ${mediaTitle}`,
      collapsed: false,
      collapsible: false,
      background: true,
    },
  ];

  for (let i = 0; i < lines.length; i += 1) {
    if (i > 0) {
      content.push({ tag: 'br' });
    }
    content.push(lines[i]!);
  }

  return [
    {
      type: 'structured-content',
      content,
    },
  ];
}

function buildSnapshotImagePath(mediaId: number, charId: number, ext: string): string {
  return `img/m${mediaId}-c${charId}.${ext}`;
}

function buildTermEntry(
  term: string,
  reading: string,
  role: CharacterDictionaryRole,
  glossary: CharacterDictionaryGlossaryEntry[],
): CharacterDictionaryTermEntry {
  const { tag, score } = roleInfo(role);
  return [term, reading, `name ${tag}`, '', score, glossary, 0, ''];
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let crc = i;
    for (let j = 0; j < 8; j += 1) {
      crc = (crc & 1) !== 0 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    table[i] = crc >>> 0;
  }
  return table;
})();

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC32_TABLE[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function createStoredZip(files: Array<{ name: string; data: Buffer }>): Buffer {
  const chunks: Buffer[] = [];
  const entries: ZipEntry[] = [];
  let offset = 0;

  for (const file of files) {
    const fileName = Buffer.from(file.name, 'utf8');
    const fileData = file.data;
    const fileCrc32 = crc32(fileData);
    const local = Buffer.alloc(30 + fileName.length);
    let cursor = 0;
    local.writeUInt32LE(0x04034b50, cursor);
    cursor += 4;
    local.writeUInt16LE(20, cursor);
    cursor += 2;
    local.writeUInt16LE(0, cursor);
    cursor += 2;
    local.writeUInt16LE(0, cursor);
    cursor += 2;
    local.writeUInt16LE(0, cursor);
    cursor += 2;
    local.writeUInt16LE(0, cursor);
    cursor += 2;
    local.writeUInt32LE(fileCrc32, cursor);
    cursor += 4;
    local.writeUInt32LE(fileData.length, cursor);
    cursor += 4;
    local.writeUInt32LE(fileData.length, cursor);
    cursor += 4;
    local.writeUInt16LE(fileName.length, cursor);
    cursor += 2;
    local.writeUInt16LE(0, cursor);
    cursor += 2;
    fileName.copy(local, cursor);

    chunks.push(local, fileData);
    entries.push({
      name: file.name,
      data: fileData,
      crc32: fileCrc32,
      localHeaderOffset: offset,
    });
    offset += local.length + fileData.length;
  }

  const centralStart = offset;
  const centralChunks: Buffer[] = [];
  for (const entry of entries) {
    const fileName = Buffer.from(entry.name, 'utf8');
    const central = Buffer.alloc(46 + fileName.length);
    let cursor = 0;
    central.writeUInt32LE(0x02014b50, cursor);
    cursor += 4;
    central.writeUInt16LE(20, cursor);
    cursor += 2;
    central.writeUInt16LE(20, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt32LE(entry.crc32, cursor);
    cursor += 4;
    central.writeUInt32LE(entry.data.length, cursor);
    cursor += 4;
    central.writeUInt32LE(entry.data.length, cursor);
    cursor += 4;
    central.writeUInt16LE(fileName.length, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt16LE(0, cursor);
    cursor += 2;
    central.writeUInt32LE(0, cursor);
    cursor += 4;
    central.writeUInt32LE(entry.localHeaderOffset, cursor);
    cursor += 4;
    fileName.copy(central, cursor);
    centralChunks.push(central);
    offset += central.length;
  }

  const centralSize = offset - centralStart;
  const end = Buffer.alloc(22);
  let cursor = 0;
  end.writeUInt32LE(0x06054b50, cursor);
  cursor += 4;
  end.writeUInt16LE(0, cursor);
  cursor += 2;
  end.writeUInt16LE(0, cursor);
  cursor += 2;
  end.writeUInt16LE(entries.length, cursor);
  cursor += 2;
  end.writeUInt16LE(entries.length, cursor);
  cursor += 2;
  end.writeUInt32LE(centralSize, cursor);
  cursor += 4;
  end.writeUInt32LE(centralStart, cursor);
  cursor += 4;
  end.writeUInt16LE(0, cursor);

  return Buffer.concat([...chunks, ...centralChunks, end]);
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

async function resolveAniListMediaIdFromGuess(
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

async function fetchCharactersForMedia(
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
                node {
                  id
                  description(asHtml: false)
                  image {
                    large
                    medium
                  }
                  name {
                    full
                    native
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
      const fullName = node.name?.full?.trim() || '';
      const nativeName = node.name?.native?.trim() || '';
      if (!fullName && !nativeName) continue;
      characters.push({
        id: node.id,
        role: mapRole(edge?.role),
        fullName,
        nativeName,
        description: node.description || '',
        imageUrl: node.image?.large || node.image?.medium || null,
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

async function downloadCharacterImage(imageUrl: string, charId: number): Promise<{
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

function buildDictionaryTitle(mediaId: number): string {
  return `SubMiner Character Dictionary (AniList ${mediaId})`;
}

function createIndex(
  dictionaryTitle: string,
  description: string,
  revision: string,
): Record<string, unknown> {
  return {
    title: dictionaryTitle,
    revision,
    format: 3,
    author: 'SubMiner',
    description,
  };
}

function createTagBank(): Array<[string, string, number, string, number]> {
  return [
    ['name', 'partOfSpeech', 0, 'Character name', 0],
    ['main', 'name', 0, 'Protagonist', 0],
    ['primary', 'name', 0, 'Main character', 0],
    ['side', 'name', 0, 'Side character', 0],
    ['appears', 'name', 0, 'Minor appearance', 0],
  ];
}

function buildSnapshotFromCharacters(
  mediaId: number,
  mediaTitle: string,
  characters: CharacterRecord[],
  imagesByCharacterId: Map<number, CharacterDictionarySnapshotImage>,
  updatedAt: number,
): CharacterDictionarySnapshot {
  const termEntries: CharacterDictionaryTermEntry[] = [];
  const seen = new Set<string>();

  for (const character of characters) {
    const imagePath = imagesByCharacterId.get(character.id)?.path ?? null;
    const glossary = createDefinitionGlossary(character, mediaTitle, imagePath);
    const candidateTerms = buildNameTerms(character);
    for (const term of candidateTerms) {
      const reading = buildReading(term);
      const dedupeKey = `${term}|${reading}|${character.role}`;
      if (seen.has(dedupeKey)) continue;
      seen.add(dedupeKey);
      termEntries.push(buildTermEntry(term, reading, character.role, glossary));
    }
  }

  if (termEntries.length === 0) {
    throw new Error('No dictionary entries generated from AniList character data.');
  }

  return {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId,
    mediaTitle,
    entryCount: termEntries.length,
    updatedAt,
    termEntries,
    images: [...imagesByCharacterId.values()],
  };
}

function buildDictionaryZip(
  outputPath: string,
  dictionaryTitle: string,
  description: string,
  revision: string,
  termEntries: CharacterDictionaryTermEntry[],
  images: CharacterDictionarySnapshotImage[],
): { zipPath: string; entryCount: number } {
  const zipFiles: Array<{ name: string; data: Buffer }> = [
    {
      name: 'index.json',
      data: Buffer.from(JSON.stringify(createIndex(dictionaryTitle, description, revision), null, 2), 'utf8'),
    },
    {
      name: 'tag_bank_1.json',
      data: Buffer.from(JSON.stringify(createTagBank()), 'utf8'),
    },
  ];

  for (const image of images) {
    zipFiles.push({
      name: image.path,
      data: Buffer.from(image.dataBase64, 'base64'),
    });
  }

  const entriesPerBank = 10_000;
  for (let i = 0; i < termEntries.length; i += entriesPerBank) {
    zipFiles.push({
      name: `term_bank_${Math.floor(i / entriesPerBank) + 1}.json`,
      data: Buffer.from(JSON.stringify(termEntries.slice(i, i + entriesPerBank)), 'utf8'),
    });
  }

  ensureDir(path.dirname(outputPath));
  fs.writeFileSync(outputPath, createStoredZip(zipFiles));
  return { zipPath: outputPath, entryCount: termEntries.length };
}

function buildMergedRevision(mediaIds: number[], snapshots: CharacterDictionarySnapshot[]): string {
  const hash = createHash('sha1');
  hash.update(
    JSON.stringify({
      mediaIds,
      snapshots: snapshots.map((snapshot) => ({
        mediaId: snapshot.mediaId,
        updatedAt: snapshot.updatedAt,
        entryCount: snapshot.entryCount,
      })),
    }),
  );
  return hash.digest('hex').slice(0, 12);
}

export function createCharacterDictionaryRuntimeService(deps: CharacterDictionaryRuntimeDeps): {
  getOrCreateCurrentSnapshot: (targetPath?: string) => Promise<CharacterDictionarySnapshotResult>;
  buildMergedDictionary: (mediaIds: number[]) => Promise<MergedCharacterDictionaryBuildResult>;
  generateForCurrentMedia: (
    targetPath?: string,
    options?: CharacterDictionaryGenerateOptions,
  ) => Promise<CharacterDictionaryBuildResult>;
} {
  const outputDir = path.join(deps.userDataPath, 'character-dictionaries');
  const sleepMs = deps.sleep ?? sleep;

  const resolveCurrentMedia = async (
    targetPath?: string,
    beforeRequest?: () => Promise<void>,
  ): Promise<ResolvedAniListMedia> => {
    deps.logInfo?.('[dictionary] resolving current anime for character dictionary generation');
    const dictionaryTarget = targetPath?.trim() || '';
    const guessInput =
      dictionaryTarget.length > 0
        ? resolveDictionaryGuessInputs(dictionaryTarget)
        : {
            mediaPath: deps.getCurrentMediaPath(),
            mediaTitle: deps.getCurrentMediaTitle(),
          };
    const mediaPathForGuess = deps.resolveMediaPathForJimaku(guessInput.mediaPath);
    const mediaTitle = guessInput.mediaTitle;
    const guessed = await deps.guessAnilistMediaInfo(mediaPathForGuess, mediaTitle);
    if (!guessed || !guessed.title.trim()) {
      throw new Error('Unable to resolve current anime from media path/title.');
    }
    deps.logInfo?.(
      `[dictionary] current anime guess: ${guessed.title.trim()}${
        typeof guessed.episode === 'number' && guessed.episode > 0 ? ` (episode ${guessed.episode})` : ''
      }`,
    );
    const resolved = await resolveAniListMediaIdFromGuess(guessed, beforeRequest);
    deps.logInfo?.(`[dictionary] AniList match: ${resolved.title} -> AniList ${resolved.id}`);
    return resolved;
  };

  const getOrCreateSnapshot = async (
    mediaId: number,
    mediaTitleHint?: string,
    beforeRequest?: () => Promise<void>,
  ): Promise<CharacterDictionarySnapshotResult> => {
    const snapshotPath = getSnapshotPath(outputDir, mediaId);
    const cachedSnapshot = readSnapshot(snapshotPath);
    if (cachedSnapshot) {
      deps.logInfo?.(`[dictionary] snapshot hit for AniList ${mediaId}`);
      return {
        mediaId: cachedSnapshot.mediaId,
        mediaTitle: cachedSnapshot.mediaTitle,
        entryCount: cachedSnapshot.entryCount,
        fromCache: true,
        updatedAt: cachedSnapshot.updatedAt,
      };
    }

    deps.logInfo?.(`[dictionary] snapshot miss for AniList ${mediaId}, fetching characters`);

    const { mediaTitle: fetchedMediaTitle, characters } = await fetchCharactersForMedia(
      mediaId,
      beforeRequest,
      (page) => {
        deps.logInfo?.(`[dictionary] downloaded AniList character page ${page} for AniList ${mediaId}`);
      },
    );
    if (characters.length === 0) {
      throw new Error(`No characters returned for AniList media ${mediaId}.`);
    }

    const imagesByCharacterId = new Map<number, CharacterDictionarySnapshotImage>();
    const charactersWithImages = characters.filter((character) => Boolean(character.imageUrl)).length;
    if (charactersWithImages > 0) {
      deps.logInfo?.(
        `[dictionary] downloading ${charactersWithImages} character images for AniList ${mediaId}`,
      );
    }
    let hasAttemptedCharacterImageDownload = false;
    for (const character of characters) {
      if (!character.imageUrl) continue;
      if (hasAttemptedCharacterImageDownload) {
        await sleepMs(CHARACTER_IMAGE_DOWNLOAD_DELAY_MS);
      }
      hasAttemptedCharacterImageDownload = true;
      const image = await downloadCharacterImage(character.imageUrl, character.id);
      if (!image) continue;
      imagesByCharacterId.set(character.id, {
        path: buildSnapshotImagePath(mediaId, character.id, image.ext),
        dataBase64: image.bytes.toString('base64'),
      });
    }

    const snapshot = buildSnapshotFromCharacters(
      mediaId,
      fetchedMediaTitle || mediaTitleHint || `AniList ${mediaId}`,
      characters,
      imagesByCharacterId,
      deps.now(),
    );
    writeSnapshot(snapshotPath, snapshot);
    deps.logInfo?.(
      `[dictionary] stored snapshot for AniList ${mediaId}: ${snapshot.entryCount} terms`,
    );

    return {
      mediaId: snapshot.mediaId,
      mediaTitle: snapshot.mediaTitle,
      entryCount: snapshot.entryCount,
      fromCache: false,
      updatedAt: snapshot.updatedAt,
    };
  };

  return {
    getOrCreateCurrentSnapshot: async (targetPath?: string) => {
      let hasAniListRequest = false;
      const waitForAniListRequestSlot = async (): Promise<void> => {
        if (!hasAniListRequest) {
          hasAniListRequest = true;
          return;
        }
        await sleepMs(ANILIST_REQUEST_DELAY_MS);
      };
      const resolvedMedia = await resolveCurrentMedia(targetPath, waitForAniListRequestSlot);
      return getOrCreateSnapshot(resolvedMedia.id, resolvedMedia.title, waitForAniListRequestSlot);
    },
    buildMergedDictionary: async (mediaIds: number[]) => {
      const normalizedMediaIds = mediaIds
        .filter((mediaId) => Number.isFinite(mediaId) && mediaId > 0)
        .map((mediaId) => Math.floor(mediaId));
      const snapshots = normalizedMediaIds.map((mediaId) => {
        const snapshot = readSnapshot(getSnapshotPath(outputDir, mediaId));
        if (!snapshot) {
          throw new Error(`Missing character dictionary snapshot for AniList ${mediaId}.`);
        }
        return snapshot;
      });
      const revision = buildMergedRevision(normalizedMediaIds, snapshots);
      const description =
        snapshots.length === 1
          ? `Character names from ${snapshots[0]!.mediaTitle}`
          : `Character names from ${snapshots.length} recent anime`;
      const { zipPath, entryCount } = buildDictionaryZip(
        getMergedZipPath(outputDir),
        CHARACTER_DICTIONARY_MERGED_TITLE,
        description,
        revision,
        snapshots.flatMap((snapshot) => snapshot.termEntries),
        snapshots.flatMap((snapshot) => snapshot.images),
      );
      deps.logInfo?.(
        `[dictionary] rebuilt merged dictionary: ${normalizedMediaIds.join(', ') || '<empty>'} -> ${zipPath}`,
      );
      return {
        zipPath,
        revision,
        dictionaryTitle: CHARACTER_DICTIONARY_MERGED_TITLE,
        entryCount,
      };
    },
    generateForCurrentMedia: async (targetPath?: string, _options?: CharacterDictionaryGenerateOptions) => {
      let hasAniListRequest = false;
      const waitForAniListRequestSlot = async (): Promise<void> => {
        if (!hasAniListRequest) {
          hasAniListRequest = true;
          return;
        }
        await sleepMs(ANILIST_REQUEST_DELAY_MS);
      };
      const resolvedMedia = await resolveCurrentMedia(targetPath, waitForAniListRequestSlot);
      const snapshot = await getOrCreateSnapshot(
        resolvedMedia.id,
        resolvedMedia.title,
        waitForAniListRequestSlot,
      );
      const storedSnapshot = readSnapshot(getSnapshotPath(outputDir, resolvedMedia.id));
      if (!storedSnapshot) {
        throw new Error(`Snapshot missing after generation for AniList ${resolvedMedia.id}.`);
      }
      const revision = String(storedSnapshot.updatedAt);
      const dictionaryTitle = buildDictionaryTitle(resolvedMedia.id);
      const description = `Character names from ${storedSnapshot.mediaTitle} [AniList media ID ${resolvedMedia.id}]`;
      const zipPath = path.join(outputDir, `anilist-${resolvedMedia.id}.zip`);
      deps.logInfo?.(`[dictionary] building ZIP for AniList ${resolvedMedia.id}`);
      buildDictionaryZip(
        zipPath,
        dictionaryTitle,
        description,
        revision,
        storedSnapshot.termEntries,
        storedSnapshot.images,
      );
      deps.logInfo?.(
        `[dictionary] generated AniList ${resolvedMedia.id}: ${storedSnapshot.entryCount} terms -> ${zipPath}`,
      );
      return {
        zipPath,
        fromCache: snapshot.fromCache,
        mediaId: resolvedMedia.id,
        mediaTitle: storedSnapshot.mediaTitle,
        entryCount: storedSnapshot.entryCount,
        dictionaryTitle,
        revision,
      };
    },
  };
}
