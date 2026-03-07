import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

import type { MergedToken } from '../types';
import { PartOfSpeech } from '../types.js';
import {
  alignTokensToSourceText,
  buildSubtitleTokenHoverRanges,
  computeWordClass,
  createSubtitleRenderer,
  getFrequencyRankLabelForToken,
  getJlptLevelLabelForToken,
  normalizeSubtitle,
  sanitizeSubtitleHoverTokenColor,
  shouldRenderTokenizedSubtitle,
} from './subtitle-render.js';
import { createRendererState } from './state.js';

class FakeTextNode {
  constructor(public textContent: string) {}
}

class FakeDocumentFragment {
  childNodes: Array<FakeElement | FakeTextNode> = [];

  appendChild(
    child: FakeElement | FakeTextNode | FakeDocumentFragment,
  ): FakeElement | FakeTextNode | FakeDocumentFragment {
    if (child instanceof FakeDocumentFragment) {
      this.childNodes.push(...child.childNodes);
      child.childNodes = [];
      return child;
    }

    this.childNodes.push(child);
    return child;
  }
}

class FakeStyleDeclaration {
  private values = new Map<string, string>();

  setProperty(name: string, value: string) {
    this.values.set(name, value);
  }
}

class FakeElement {
  childNodes: Array<FakeElement | FakeTextNode> = [];
  dataset: Record<string, string> = {};
  style = new FakeStyleDeclaration();
  className = '';
  private ownTextContent = '';

  constructor(public tagName: string) {}

  appendChild(
    child: FakeElement | FakeTextNode | FakeDocumentFragment,
  ): FakeElement | FakeTextNode | FakeDocumentFragment {
    if (child instanceof FakeDocumentFragment) {
      this.childNodes.push(...child.childNodes);
      child.childNodes = [];
      return child;
    }

    this.childNodes.push(child);
    return child;
  }

  set textContent(value: string) {
    this.ownTextContent = value;
    this.childNodes = [];
  }

  get textContent(): string {
    if (this.childNodes.length === 0) {
      return this.ownTextContent;
    }

    return this.childNodes
      .map((child) => (child instanceof FakeTextNode ? child.textContent : child.textContent))
      .join('');
  }

  set innerHTML(value: string) {
    if (value === '') {
      this.childNodes = [];
      this.ownTextContent = '';
    }
  }
}

function installFakeDocument() {
  const previousDocument = (globalThis as { document?: unknown }).document;

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createDocumentFragment: () => new FakeDocumentFragment(),
      createElement: (tagName: string) => new FakeElement(tagName),
      createTextNode: (text: string) => new FakeTextNode(text),
    },
  });

  return () => {
    Object.defineProperty(globalThis, 'document', {
      configurable: true,
      value: previousDocument,
    });
  };
}

function collectWordNodes(root: FakeElement): FakeElement[] {
  return root.childNodes.filter(
    (child): child is FakeElement => child instanceof FakeElement && child.className.includes('word'),
  );
}

function createToken(overrides: Partial<MergedToken>): MergedToken {
  return {
    surface: '',
    reading: '',
    headword: '',
    startPos: 0,
    endPos: 0,
    partOfSpeech: PartOfSpeech.other,
    isMerged: true,
    isKnown: false,
    isNPlusOneTarget: false,
    ...overrides,
  };
}

function extractClassBlock(cssText: string, selector: string): string {
  const ruleRegex = /([^{}]+)\{([^}]*)\}/g;
  let match: RegExpExecArray | null = null;
  let fallbackBlock = '';

  while ((match = ruleRegex.exec(cssText)) !== null) {
    const selectorsBlock = match[1]?.trim() ?? '';
    const selectorBlock = match[2] ?? '';

    const selectors = selectorsBlock
      .split(',')
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);

    if (selectors.includes(selector)) {
      if (selectors.length === 1) {
        return selectorBlock;
      }

      if (!fallbackBlock) {
        fallbackBlock = selectorBlock;
      }
    }
  }

  if (fallbackBlock) {
    return fallbackBlock;
  }

  return '';
}

test('computeWordClass preserves known and n+1 classes while adding JLPT classes', () => {
  const knownJlpt = createToken({
    isKnown: true,
    jlptLevel: 'N1',
    surface: '猫',
  });
  const nPlusOneJlpt = createToken({
    isNPlusOneTarget: true,
    jlptLevel: 'N2',
    surface: '犬',
  });

  assert.equal(computeWordClass(knownJlpt), 'word word-known word-jlpt-n1');
  assert.equal(computeWordClass(nPlusOneJlpt), 'word word-n-plus-one word-jlpt-n2');
});

test('computeWordClass applies name-match class ahead of known and frequency classes', () => {
  const token = createToken({
    isKnown: true,
    frequencyRank: 10,
    surface: 'アクア',
  }) as MergedToken & { isNameMatch?: boolean };
  token.isNameMatch = true;

  assert.equal(
    computeWordClass(token, {
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-name-match',
  );
});

test('computeWordClass skips name-match class when disabled', () => {
  const token = createToken({
    surface: 'アクア',
  }) as MergedToken & { isNameMatch?: boolean };
  token.isNameMatch = true;

  assert.equal(
    computeWordClass(token, {
      nameMatchEnabled: false,
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word',
  );
});

test('computeWordClass keeps known and N+1 color classes exclusive over frequency classes', () => {
  const known = createToken({
    isKnown: true,
    frequencyRank: 10,
    surface: '既知',
  });
  const nPlusOne = createToken({
    isNPlusOneTarget: true,
    frequencyRank: 10,
    surface: '目標',
  });
  const frequency = createToken({
    frequencyRank: 10,
    surface: '頻度',
  });

  assert.equal(
    computeWordClass(known, {
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-known',
  );
  assert.equal(
    computeWordClass(nPlusOne, {
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-n-plus-one',
  );
  assert.equal(
    computeWordClass(frequency, {
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-frequency-single',
  );
});

test('applySubtitleStyle sets subtitle name-match color variable', () => {
  const restoreDocument = installFakeDocument();
  try {
    const subtitleRoot = new FakeElement('div');
    const subtitleContainer = new FakeElement('div');
    const secondarySubRoot = new FakeElement('div');
    const secondarySubContainer = new FakeElement('div');
    const ctx = {
      state: createRendererState(),
      dom: {
        subtitleRoot,
        subtitleContainer,
        secondarySubRoot,
        secondarySubContainer,
      },
    } as never;

    const renderer = createSubtitleRenderer(ctx);
    renderer.applySubtitleStyle({
      nameMatchColor: '#f5bde6',
    } as never);

    assert.equal(
      (subtitleRoot.style as unknown as { values?: Map<string, string> }).values?.get(
        '--subtitle-name-match-color',
      ),
      '#f5bde6',
    );
  } finally {
    restoreDocument();
  }
});

test('computeWordClass adds frequency class for single mode when rank is within topX', () => {
  const token = createToken({
    surface: '猫',
    frequencyRank: 50,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 100,
    mode: 'single',
    singleColor: '#000000',
    bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
  });

  assert.equal(actual, 'word word-frequency-single');
});

test('computeWordClass adds frequency class when rank equals topX', () => {
  const token = createToken({
    surface: '水',
    frequencyRank: 100,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 100,
    mode: 'single',
    singleColor: '#000000',
    bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
  });

  assert.equal(actual, 'word word-frequency-single');
});

test('computeWordClass adds frequency class for banded mode', () => {
  const token = createToken({
    surface: '犬',
    frequencyRank: 250,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 1000,
    mode: 'banded',
    singleColor: '#000000',
    bandedColors: ['#111111', '#222222', '#333333', '#444444', '#555555'] as const,
  });

  assert.equal(actual, 'word word-frequency-band-2');
});

test('computeWordClass uses configured band count for banded mode', () => {
  const token = createToken({
    surface: '犬',
    frequencyRank: 2,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 4,
    mode: 'banded',
    singleColor: '#000000',
    bandedColors: ['#111111', '#222222', '#333333', '#444444', '#555555'],
  } as any);

  assert.equal(actual, 'word word-frequency-band-3');
});

test('computeWordClass skips frequency class when rank is out of topX', () => {
  const token = createToken({
    surface: '犬',
    frequencyRank: 1200,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 1000,
    mode: 'single',
    singleColor: '#000000',
    bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
  });

  assert.equal(actual, 'word');
});

test('getFrequencyRankLabelForToken returns rank only for frequency-colored tokens', () => {
  const settings = {
    enabled: true,
    topX: 100,
    mode: 'single' as const,
    singleColor: '#000000',
    bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as [
      string,
      string,
      string,
      string,
      string,
    ],
  };
  const frequencyToken = createToken({ surface: '頻度', frequencyRank: 20 });
  const knownToken = createToken({ surface: '既知', isKnown: true, frequencyRank: 20 });
  const nPlusOneToken = createToken({ surface: '目標', isNPlusOneTarget: true, frequencyRank: 20 });
  const outOfRangeToken = createToken({ surface: '圏外', frequencyRank: 1000 });

  assert.equal(getFrequencyRankLabelForToken(frequencyToken, settings), '20');
  assert.equal(getFrequencyRankLabelForToken(knownToken, settings), '20');
  assert.equal(getFrequencyRankLabelForToken(nPlusOneToken, settings), '20');
  assert.equal(getFrequencyRankLabelForToken(outOfRangeToken, settings), null);
});

test('getJlptLevelLabelForToken returns level when token has jlpt metadata', () => {
  const jlptToken = createToken({ surface: '語彙', jlptLevel: 'N2' });
  const noJlptToken = createToken({ surface: '語彙' });

  assert.equal(getJlptLevelLabelForToken(jlptToken), 'N2');
  assert.equal(getJlptLevelLabelForToken(noJlptToken), null);
});

test('sanitizeSubtitleHoverTokenColor falls back for pure black values', () => {
  assert.equal(sanitizeSubtitleHoverTokenColor('#000000'), '#f4dbd6');
  assert.equal(sanitizeSubtitleHoverTokenColor('000000'), '#f4dbd6');
  assert.equal(sanitizeSubtitleHoverTokenColor('#0000'), '#f4dbd6');
});

test('sanitizeSubtitleHoverTokenColor keeps non-black color values', () => {
  assert.equal(sanitizeSubtitleHoverTokenColor('#ff00ff'), '#ff00ff');
  assert.equal(sanitizeSubtitleHoverTokenColor(undefined), '#f4dbd6');
});

test('alignTokensToSourceText preserves newline separators between adjacent token surfaces', () => {
  const tokens = [
    createToken({ surface: 'キリキリと', reading: 'きりきりと', headword: 'キリキリと' }),
    createToken({ surface: 'かかってこい', reading: 'かかってこい', headword: 'かかってこい' }),
  ];

  const segments = alignTokensToSourceText(tokens, 'キリキリと\nかかってこい');
  assert.deepEqual(
    segments.map((segment) => (segment.kind === 'text' ? `text:${segment.text}` : 'token')),
    ['token', 'text:\n', 'token'],
  );
});

test('alignTokensToSourceText treats whitespace-only token surfaces as plain text separators', () => {
  const tokens = [
    createToken({ surface: '常人が使えば' }),
    createToken({ surface: ' ' }),
    createToken({ surface: 'その圧倒的な力に' }),
    createToken({ surface: '\n' }),
    createToken({ surface: '体が耐えきれず死に至るが…' }),
  ];

  const segments = alignTokensToSourceText(
    tokens,
    '常人が使えば その圧倒的な力に\n体が耐えきれず死に至るが…',
  );
  assert.deepEqual(
    segments.map((segment) => (segment.kind === 'text' ? `text:${segment.text}` : 'token')),
    ['token', 'text: ', 'token', 'text:\n', 'token'],
  );
});

test('alignTokensToSourceText preserves unsupported punctuation between matched tokens', () => {
  const tokens = [createToken({ surface: 'えっ' }), createToken({ surface: 'マジ' })];

  const segments = alignTokensToSourceText(tokens, 'えっ！？マジ');
  assert.deepEqual(
    segments.map((segment) => (segment.kind === 'text' ? `text:${segment.text}` : 'token')),
    ['token', 'text:！？', 'token'],
  );
});

test('alignTokensToSourceText avoids duplicate tail when later token surface does not match source', () => {
  const tokens = [
    createToken({ surface: '君たちが潰した拠点に' }),
    createToken({ surface: '教団の主力は1人もいない' }),
  ];

  const segments = alignTokensToSourceText(
    tokens,
    '君たちが潰した拠点に\n教団の主力は１人もいない',
  );
  assert.deepEqual(
    segments.map((segment) => (segment.kind === 'text' ? `text:${segment.text}` : 'token')),
    ['token', 'text:\n教団の主力は１人もいない'],
  );
});

test('buildSubtitleTokenHoverRanges tracks token offsets across text separators', () => {
  const tokens = [createToken({ surface: 'キリキリと' }), createToken({ surface: 'かかってこい' })];

  const ranges = buildSubtitleTokenHoverRanges(tokens, 'キリキリと\nかかってこい');
  assert.deepEqual(ranges, [
    { start: 0, end: 5, tokenIndex: 0 },
    { start: 6, end: 12, tokenIndex: 1 },
  ]);
});

test('buildSubtitleTokenHoverRanges ignores unmatched token surfaces', () => {
  const tokens = [
    createToken({ surface: '君たちが潰した拠点に' }),
    createToken({ surface: '教団の主力は1人もいない' }),
  ];

  const ranges = buildSubtitleTokenHoverRanges(
    tokens,
    '君たちが潰した拠点に\n教団の主力は１人もいない',
  );
  assert.deepEqual(ranges, [{ start: 0, end: 10, tokenIndex: 0 }]);
});

test('buildSubtitleTokenHoverRanges skips unsupported punctuation while preserving later offsets', () => {
  const tokens = [createToken({ surface: 'えっ' }), createToken({ surface: 'マジ' })];

  const ranges = buildSubtitleTokenHoverRanges(tokens, 'えっ！？マジ');
  assert.deepEqual(ranges, [
    { start: 0, end: 2, tokenIndex: 0 },
    { start: 4, end: 6, tokenIndex: 1 },
  ]);
});

test('renderSubtitle preserves unsupported punctuation while keeping it non-interactive', () => {
  const restoreDocument = installFakeDocument();

  try {
    const subtitleRoot = new FakeElement('div');
    const renderer = createSubtitleRenderer({
      dom: {
        subtitleRoot,
        subtitleContainer: new FakeElement('div'),
        secondarySubRoot: new FakeElement('div'),
        secondarySubContainer: new FakeElement('div'),
      },
      platform: {
        isMacOSPlatform: false,
        isModalLayer: false,
        overlayLayer: 'visible',
        shouldToggleMouseIgnore: false,
      },
      state: createRendererState(),
    } as never);

    renderer.renderSubtitle({
      text: 'えっ！？マジ',
      tokens: [createToken({ surface: 'えっ' }), createToken({ surface: 'マジ' })],
    });

    assert.equal(subtitleRoot.textContent, 'えっ！？マジ');
    assert.deepEqual(
      collectWordNodes(subtitleRoot).map((node) => [node.textContent, node.dataset.tokenIndex]),
      [
        ['えっ', '0'],
        ['マジ', '1'],
      ],
    );
  } finally {
    restoreDocument();
  }
});

test('normalizeSubtitle collapses explicit line breaks when collapseLineBreaks is enabled', () => {
  assert.equal(
    normalizeSubtitle('常人が使えば\\Nその圧倒的な力に\\n体が耐えきれず死に至るが…', true, true),
    '常人が使えば その圧倒的な力に 体が耐えきれず死に至るが…',
  );
});

test('shouldRenderTokenizedSubtitle enables token rendering when tokens exist', () => {
  assert.equal(shouldRenderTokenizedSubtitle(5), true);
  assert.equal(shouldRenderTokenizedSubtitle(0), false);
});

test('JLPT CSS rules use underline-only styling in renderer stylesheet', () => {
  const distCssPath = path.join(process.cwd(), 'dist', 'renderer', 'style.css');
  const srcCssPath = path.join(process.cwd(), 'src', 'renderer', 'style.css');

  const cssPath = fs.existsSync(srcCssPath) ? srcCssPath : distCssPath;
  if (!fs.existsSync(cssPath)) {
    assert.fail(
      'JLPT CSS file missing. Run `bun run build` first, or ensure src/renderer/style.css exists.',
    );
  }

  const cssText = fs.readFileSync(cssPath, 'utf-8');

  for (let level = 1; level <= 5; level += 1) {
    const block = extractClassBlock(cssText, `#subtitleRoot .word.word-jlpt-n${level}`);
    assert.ok(block.length > 0, `word-jlpt-n${level} class should exist`);
    assert.match(block, /text-decoration-line:\s*underline;/);
    assert.match(block, /text-decoration-thickness:\s*2px;/);
    assert.match(block, /text-underline-offset:\s*4px;/);
    assert.doesNotMatch(block, /(?:^|\n)\s*color\s*:/m);
  }

  for (let band = 1; band <= 5; band += 1) {
    const block = extractClassBlock(
      cssText,
      band === 1
        ? '#subtitleRoot .word.word-frequency-single'
        : `#subtitleRoot .word.word-frequency-band-${band}`,
    );
    assert.ok(
      block.length > 0,
      `frequency class word-frequency-${band === 1 ? 'single' : `band-${band}`} should exist`,
    );
    assert.match(block, /color:\s*var\(/);
  }

  const visibleMacBlock = extractClassBlock(
    cssText,
    'body.platform-macos.layer-visible #subtitleRoot',
  );
  assert.match(visibleMacBlock, /--visible-sub-line-height:\s*1\.64;/);
  assert.match(visibleMacBlock, /--visible-sub-line-gap:\s*0\.54em;/);

  const subtitleRootBlock = extractClassBlock(cssText, '#subtitleRoot');
  assert.match(subtitleRootBlock, /--subtitle-hover-token-color:\s*#f4dbd6;/);
  assert.match(
    subtitleRootBlock,
    /--subtitle-hover-token-background-color:\s*rgba\(54,\s*58,\s*79,\s*0\.84\);/,
  );
  assert.match(subtitleRootBlock, /-webkit-text-fill-color:\s*currentColor;/);

  const charBlock = extractClassBlock(cssText, '#subtitleRoot .c');
  assert.match(charBlock, /-webkit-text-fill-color:\s*currentColor\s*!important;/);

  const wordBlock = extractClassBlock(cssText, '#subtitleRoot .word');
  assert.match(wordBlock, /-webkit-text-fill-color:\s*currentColor\s*!important;/);

  const frequencyTooltipBaseBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word[data-frequency-rank]::before',
  );
  assert.match(frequencyTooltipBaseBlock, /content:\s*attr\(data-frequency-rank\);/);
  assert.match(frequencyTooltipBaseBlock, /opacity:\s*0;/);
  assert.match(frequencyTooltipBaseBlock, /pointer-events:\s*none;/);

  const frequencyTooltipHoverBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word[data-frequency-rank]:hover::before',
  );
  assert.match(frequencyTooltipHoverBlock, /opacity:\s*1;/);
  const frequencyTooltipKeyboardSelectedBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word.keyboard-selected[data-frequency-rank]::before',
  );
  assert.match(frequencyTooltipKeyboardSelectedBlock, /opacity:\s*1;/);

  const jlptTooltipBaseBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word[data-jlpt-level]::after',
  );
  assert.match(jlptTooltipBaseBlock, /content:\s*attr\(data-jlpt-level\);/);
  assert.match(jlptTooltipBaseBlock, /bottom:\s*-\s*0\.42em;/);
  assert.match(jlptTooltipBaseBlock, /opacity:\s*0;/);
  assert.match(jlptTooltipBaseBlock, /pointer-events:\s*none;/);

  const jlptTooltipHoverBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word[data-jlpt-level]:hover::after',
  );
  assert.match(jlptTooltipHoverBlock, /opacity:\s*1;/);
  const jlptTooltipKeyboardSelectedBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word.keyboard-selected[data-jlpt-level]::after',
  );
  assert.match(jlptTooltipKeyboardSelectedBlock, /opacity:\s*1;/);

  assert.match(
    cssText,
    /#subtitleRoot\s+\.word:not\(\.word-known\):not\(\.word-n-plus-one\):not\(\.word-name-match\):not\(\.word-frequency-single\):not\(\s*\.word-frequency-band-1\s*\):not\(\.word-frequency-band-2\):not\(\.word-frequency-band-3\):not\(\.word-frequency-band-4\):not\(\s*\.word-frequency-band-5\s*\):hover\s*\{[\s\S]*?background:\s*var\(--subtitle-hover-token-background-color,\s*rgba\(54,\s*58,\s*79,\s*0\.84\)\);[\s\S]*?color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;[\s\S]*?-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  const coloredWordHoverBlock = extractClassBlock(cssText, '#subtitleRoot .word.word-known:hover');
  assert.match(
    coloredWordHoverBlock,
    /background:\s*var\(--subtitle-hover-token-background-color,\s*rgba\(54,\s*58,\s*79,\s*0\.84\)\);/,
  );
  assert.match(coloredWordHoverBlock, /border-radius:\s*3px;/);
  assert.match(coloredWordHoverBlock, /font-weight:\s*800;/);
  assert.doesNotMatch(coloredWordHoverBlock, /color:\s*var\(--subtitle-hover-token-color/);
  assert.doesNotMatch(
    coloredWordHoverBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color/,
  );

  const coloredWordSelectionBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word.word-known::selection',
  );
  assert.match(
    coloredWordSelectionBlock,
    /color:\s*var\(--subtitle-known-word-color,\s*#a6da95\)\s*!important;/,
  );
  assert.match(
    coloredWordSelectionBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-known-word-color,\s*#a6da95\)\s*!important;/,
  );

  const coloredCharHoverBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word.word-known .c:hover',
  );
  assert.match(coloredCharHoverBlock, /background:\s*transparent;/);
  assert.match(coloredCharHoverBlock, /color:\s*inherit\s*!important;/);

  assert.match(
    cssText,
    /\.word:is\(\.word-jlpt-n1,\s*\.word-jlpt-n2,\s*\.word-jlpt-n3,\s*\.word-jlpt-n4,\s*\.word-jlpt-n5\):not\(\s*\.word-known\s*\):not\(\.word-n-plus-one\):not\(\.word-name-match\):not\(\.word-frequency-single\):not\(\.word-frequency-band-1\):not\(\s*\.word-frequency-band-2\s*\):not\(\.word-frequency-band-3\):not\(\.word-frequency-band-4\):not\(\.word-frequency-band-5\):hover\s*\{[\s\S]*?color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;[\s\S]*?-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );
  assert.match(
    cssText,
    /\.word:is\(\.word-jlpt-n1,\s*\.word-jlpt-n2,\s*\.word-jlpt-n3,\s*\.word-jlpt-n4,\s*\.word-jlpt-n5\):not\(\s*\.word-known\s*\):not\(\.word-n-plus-one\):not\(\.word-name-match\):not\(\.word-frequency-single\):not\(\.word-frequency-band-1\):not\(\s*\.word-frequency-band-2\s*\):not\(\.word-frequency-band-3\):not\(\.word-frequency-band-4\):not\(\.word-frequency-band-5\)::selection[\s\S]*?color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;[\s\S]*?-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  const selectionBlock = extractClassBlock(cssText, '#subtitleRoot::selection');
  assert.match(
    selectionBlock,
    /background:\s*var\(--subtitle-hover-token-background-color,\s*rgba\(54,\s*58,\s*79,\s*0\.84\)\);/,
  );
  assert.match(
    selectionBlock,
    /color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );
  assert.match(
    selectionBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  const descendantSelectionBlock = extractClassBlock(cssText, '#subtitleRoot *::selection');
  assert.match(
    descendantSelectionBlock,
    /background:\s*var\(--subtitle-hover-token-background-color,\s*rgba\(54,\s*58,\s*79,\s*0\.84\)\)\s*!important;/,
  );
  assert.match(
    descendantSelectionBlock,
    /color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );
  assert.match(
    descendantSelectionBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  assert.doesNotMatch(
    cssText,
    /body\.layer-visible\s+#secondarySubContainer\s*\{[^}]*display:\s*none/i,
  );
});
