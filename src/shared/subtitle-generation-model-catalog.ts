// Official multilingual models from whisper.cpp/models/download-ggml-model.sh.
// Byte sizes and SHA256: https://huggingface.co/api/models/ggerganov/whisper.cpp/tree/main.
const MODEL_CATALOG = {
  tiny: {
    id: 'tiny',
    size: 77691713,
    sha256: 'be07e048e1e599ad46341c8d2a135645097a538221678b7acdd1b1919c6e1b21',
    description: 'Fastest and lightest; more transcription errors.',
  },
  'tiny-q5_1': {
    id: 'tiny-q5_1',
    size: 32152673,
    sha256: '818710568da3ca15689e31a743197b520007872ff9576237bda97bd1b469c3d7',
    description:
      'tiny with 5-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'tiny-q8_0': {
    id: 'tiny-q8_0',
    size: 43537433,
    sha256: 'c2085835d3f50733e2ff6e4b41ae8a2b8d8110461e18821b09a15c40c42d1cca',
    description:
      'tiny with 8-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  base: {
    id: 'base',
    size: 147951465,
    sha256: '60ed5bc3dd14eea856493d334349b405782ddcaf0028d4b5df4088345fba2efe',
    description: 'Fast and lightweight; less accurate than small.',
  },
  'base-q5_1': {
    id: 'base-q5_1',
    size: 59707625,
    sha256: '422f1ae452ade6f30a004d7e5c6a43195e4433bc370bf23fac9cc591f01a8898',
    description:
      'base with 5-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'base-q8_0': {
    id: 'base-q8_0',
    size: 81768585,
    sha256: 'c577b9a86e7e048a0b7eada054f4dd79a56bbfa911fbdacf900ac5b567cbb7d9',
    description:
      'base with 8-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  small: {
    id: 'small',
    size: 487601967,
    sha256: '1be3a9b2063867b937e64e2ec7483364a79917e157fa98c5d94b5c1fffea987b',
    description: 'Recommended starting point for accuracy and processing time.',
  },
  'small-q5_1': {
    id: 'small-q5_1',
    size: 190085487,
    sha256: 'ae85e4a935d7a567bd102fe55afc16bb595bdb618e11b2fc7591bc08120411bb',
    description:
      'small with 5-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'small-q8_0': {
    id: 'small-q8_0',
    size: 264464607,
    sha256: '49c8fb02b65e6049d5fa6c04f81f53b867b5ec9540406812c643f177317f779f',
    description:
      'small with 8-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  medium: {
    id: 'medium',
    size: 1533763059,
    sha256: '6c14d5adee5f86394037b4e4e8b59f1673b6cee10e3cf0b11bbdbee79c156208',
    description: 'Prioritizes accuracy over small; takes longer to process.',
  },
  'medium-q5_0': {
    id: 'medium-q5_0',
    size: 539212467,
    sha256: '19fea4b380c3a618ec4723c3eef2eb785ffba0d0538cf43f8f235e7b3b34220f',
    description:
      'medium with 5-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'medium-q8_0': {
    id: 'medium-q8_0',
    size: 823369779,
    sha256: '42a1ffcbe4167d224232443396968db4d02d4e8e87e213d3ee2e03095dea6502',
    description:
      'medium with 8-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'large-v1': {
    id: 'large-v1',
    size: 3094623691,
    sha256: '7d99f41a10525d0206bddadd86760181fa920438b6b33237e3118ff6c83bb53d',
    description: 'Older large model; high memory use and longer processing.',
  },
  'large-v2': {
    id: 'large-v2',
    size: 3094623691,
    sha256: '9a423fe4d40c82774b6af34115b8b935f34152246eb19e80e376071d3f999487',
    description: 'Older large model; high memory use and longer processing.',
  },
  'large-v2-q5_0': {
    id: 'large-v2-q5_0',
    size: 1080732091,
    sha256: '3a214837221e4530dbc1fe8d734f302af393eb30bd0ed046042ebf4baf70f6f2',
    description:
      'large-v2 with 5-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'large-v2-q8_0': {
    id: 'large-v2-q8_0',
    size: 1656129691,
    sha256: 'fef54e6d898246a65c8285bfa83bd1807e27fadf54d5d4e81754c47634737e8c',
    description:
      'large-v2 with 8-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'large-v3': {
    id: 'large-v3',
    size: 3095033483,
    sha256: '64d182b440b98d5203c4f9bd541544d84c605196c4f7b845dfa11fb23594d1e2',
    description: 'Prioritizes accuracy; high memory use and longer processing.',
  },
  'large-v3-q5_0': {
    id: 'large-v3-q5_0',
    size: 1081140203,
    sha256: 'd75795ecff3f83b5faa89d1900604ad8c780abd5739fae406de19f23ecd98ad1',
    description:
      'large-v3 with 5-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'large-v3-turbo': {
    id: 'large-v3-turbo',
    size: 1624555275,
    sha256: '1fc70f774d38eb169993ac391eea357ef47c88757ef72ee5943879b7e8e2bc69',
    description: 'Faster large-v3 variant with an accuracy tradeoff; uses more memory than small.',
  },
  'large-v3-turbo-q5_0': {
    id: 'large-v3-turbo-q5_0',
    size: 574041195,
    sha256: '394221709cd5ad1f40c46e6031ca61bce88931e6e088c188294c6d5a55ffa7e2',
    description:
      'large-v3-turbo with 5-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
  'large-v3-turbo-q8_0': {
    id: 'large-v3-turbo-q8_0',
    size: 874188075,
    sha256: '317eb69c11673c9de1e1f0d459b253999804ec71ac4c23c17ecf5fbe24e259a1',
    description:
      'large-v3-turbo with 8-bit quantization: lower memory use, with a possible accuracy tradeoff.',
  },
} as const;

export type SubtitleGenerationModelId = keyof typeof MODEL_CATALOG;

export const SUBTITLE_GENERATION_MODELS = Object.values(MODEL_CATALOG);

export const RECOMMENDED_SUBTITLE_GENERATION_MODEL = 'small' satisfies SubtitleGenerationModelId;

export function isSubtitleGenerationModelId(value: unknown): value is SubtitleGenerationModelId {
  return typeof value === 'string' && Object.hasOwn(MODEL_CATALOG, value);
}

export function getSubtitleGenerationModel(id: SubtitleGenerationModelId) {
  return MODEL_CATALOG[id];
}

export function formatSubtitleGenerationModelSize(size: number): string {
  if (size < 1024 ** 2) return `${Math.ceil(size / 1024)} KiB`;
  return size >= 1024 ** 3
    ? `${(size / 1024 ** 3).toFixed(1)} GiB`
    : `${Math.round(size / 1024 ** 2)} MiB`;
}
