export const ANILIST_GRAPHQL_URL = 'https://graphql.anilist.co';
export const ANILIST_REQUEST_DELAY_MS = 2000;
export const CHARACTER_IMAGE_DOWNLOAD_DELAY_MS = 250;
export const CHARACTER_DICTIONARY_FORMAT_VERSION = 20;
export const CHARACTER_DICTIONARY_MERGED_TITLE = 'SubMiner Character Dictionary';

export const HONORIFIC_SUFFIXES = [
  { term: 'さん', reading: 'さん' },
  { term: '様', reading: 'さま' },
  { term: '先生', reading: 'せんせい' },
  { term: '先輩', reading: 'せんぱい' },
  { term: '後輩', reading: 'こうはい' },
  { term: '氏', reading: 'し' },
  { term: '君', reading: 'くん' },
  { term: 'くん', reading: 'くん' },
  { term: 'ちゃん', reading: 'ちゃん' },
  { term: 'たん', reading: 'たん' },
  { term: '坊', reading: 'ぼう' },
  { term: '殿', reading: 'どの' },
  { term: '博士', reading: 'はかせ' },
  { term: '社長', reading: 'しゃちょう' },
  { term: '部長', reading: 'ぶちょう' },
] as const;
