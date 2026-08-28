import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

import type { MergedToken } from '../types';
import {
  alignTokensToSourceText,
  buildSubtitleTokenHoverRanges,
  createSubtitleRenderer,
  getFrequencyRankLabelForToken,
  getJlptLevelLabelForToken,
  normalizeSubtitle,
  normalizeSubtitleForDisplay,
  prepareSecondarySubtitleLines,
  sanitizeSubtitleHoverTokenColor,
  shouldRenderTokenizedSubtitle,
} from './subtitle-render.js';
import { createToken } from './subtitle-render-test-helpers.js';
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

  removeProperty(name: string) {
    const previous = this.values.get(name) ?? '';
    this.values.delete(name);
    return previous;
  }
}

class FakeElement {
  childNodes: Array<FakeElement | FakeTextNode> = [];
  dataset: Record<string, string> = {};
  style = new FakeStyleDeclaration();
  className = '';
  replaceChildrenCalls = 0;
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

  replaceChildren(): void {
    this.replaceChildrenCalls += 1;
    this.childNodes = [];
    this.ownTextContent = '';
  }

  cloneNode(_deep: boolean): FakeElement {
    return new FakeElement(this.tagName);
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
    (child): child is FakeElement =>
      child instanceof FakeElement && child.className.includes('word'),
  );
}

function extractClassBlock(cssText: string, selector: string): string {
  const ruleRegex = /([^{}]+)\{([^}]*)\}/g;
  let match: RegExpExecArray | null = null;
  let fallbackBlock = '';
  const normalizedSelector = normalizeCssSelector(selector);

  while ((match = ruleRegex.exec(cssText)) !== null) {
    const selectorsBlock = match[1]?.trim() ?? '';
    const selectorBlock = match[2] ?? '';

    const selectors = splitCssSelectors(selectorsBlock);

    if (selectors.some((entry) => normalizeCssSelector(entry) === normalizedSelector)) {
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

function splitCssSelectors(selectorsBlock: string): string[] {
  const selectors: string[] = [];
  let current = '';
  let parenDepth = 0;

  for (const char of selectorsBlock) {
    if (char === '(') {
      parenDepth += 1;
      current += char;
      continue;
    }

    if (char === ')') {
      parenDepth = Math.max(0, parenDepth - 1);
      current += char;
      continue;
    }

    if (char === ',' && parenDepth === 0) {
      const trimmed = current.trim();
      if (trimmed.length > 0) {
        selectors.push(trimmed);
      }
      current = '';
      continue;
    }

    current += char;
  }

  const trimmed = current.trim();
  if (trimmed.length > 0) {
    selectors.push(trimmed);
  }

  return selectors;
}

function normalizeCssSelector(selector: string): string {
  return selector
    .replace(/\s+/g, ' ')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .replace(/\s*,\s*/g, ', ')
    .trim();
}

function buildJlptColorSelector(level: number): string {
  const higherPriorityClasses = [
    '.word-known',
    '.word-n-plus-one',
    '.word-name-match',
    '.word-frequency-single',
    '.word-frequency-band-1',
    '.word-frequency-band-2',
    '.word-frequency-band-3',
    '.word-frequency-band-4',
    '.word-frequency-band-5',
  ].join(', ');

  return `#subtitleRoot .word.word-jlpt-n${level}:not(:is(${higherPriorityClasses}))`;
}

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

test('renderSubtitle injects circular character image for annotated name matches', () => {
  const restoreDocument = installFakeDocument();
  try {
    const subtitleRoot = new FakeElement('div');
    const ctx = {
      state: {
        ...createRendererState(),
        nameMatchEnabled: true,
      },
      dom: {
        subtitleRoot,
        subtitleContainer: new FakeElement('div'),
        secondarySubRoot: new FakeElement('div'),
        secondarySubContainer: new FakeElement('div'),
      },
    } as never;

    const renderer = createSubtitleRenderer(ctx);
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [
        {
          ...createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' }),
          isNameMatch: true,
          characterImage: {
            src: 'data:image/png;base64,AAAA',
            alt: 'アクア',
          },
        } as MergedToken,
      ],
    });

    const [word] = collectWordNodes(subtitleRoot);
    assert.ok(word);
    assert.equal(word.className, 'word word-name-match word-character-image-token');
    assert.equal(word.textContent, 'アクア');
    const image = word.childNodes[0] as FakeElement & { src?: string; alt?: string };
    assert.equal(image.tagName, 'img');
    assert.equal(image.className, 'word-character-image');
    assert.equal(image.src, 'data:image/png;base64,AAAA');
    assert.equal(image.alt, 'アクア');
  } finally {
    restoreDocument();
  }
});

test('renderSubtitle skips character image when name-match rendering is disabled', () => {
  const restoreDocument = installFakeDocument();
  try {
    const subtitleRoot = new FakeElement('div');
    const ctx = {
      state: {
        ...createRendererState(),
        nameMatchEnabled: false,
      },
      dom: {
        subtitleRoot,
        subtitleContainer: new FakeElement('div'),
        secondarySubRoot: new FakeElement('div'),
        secondarySubContainer: new FakeElement('div'),
      },
    } as never;

    const renderer = createSubtitleRenderer(ctx);
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [
        {
          ...createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' }),
          isNameMatch: true,
          characterImage: {
            src: 'data:image/png;base64,AAAA',
            alt: 'アクア',
          },
        } as MergedToken,
      ],
    });

    const [word] = collectWordNodes(subtitleRoot);
    assert.ok(word);
    assert.equal(word.className, 'word');
    assert.equal(word.textContent, 'アクア');
    assert.equal(word.childNodes.length, 0);
  } finally {
    restoreDocument();
  }
});

test('renderSubtitle skips identical primary subtitle DOM replacement', () => {
  const restoreDocument = installFakeDocument();
  try {
    const subtitleRoot = new FakeElement('div');
    const ctx = {
      state: createRendererState(),
      dom: {
        subtitleRoot,
        subtitleContainer: new FakeElement('div'),
        secondarySubRoot: new FakeElement('div'),
        secondarySubContainer: new FakeElement('div'),
      },
    } as never;

    const renderer = createSubtitleRenderer(ctx);
    renderer.renderSubtitle({ text: '字幕', tokens: null });
    renderer.renderSubtitle({ text: '字幕', tokens: null });
    renderer.renderSubtitle({ text: '字幕2', tokens: null });

    assert.equal(subtitleRoot.replaceChildrenCalls, 2);
    assert.equal(subtitleRoot.textContent, '字幕2');
  } finally {
    restoreDocument();
  }
});

test('renderSubtitle keeps tokenized subtitle when stale plain payload repeats same text', () => {
  const restoreDocument = installFakeDocument();
  try {
    const subtitleRoot = new FakeElement('div');
    const ctx = {
      state: createRendererState(),
      dom: {
        subtitleRoot,
        subtitleContainer: new FakeElement('div'),
        secondarySubRoot: new FakeElement('div'),
        secondarySubContainer: new FakeElement('div'),
      },
    } as never;

    const renderer = createSubtitleRenderer(ctx);
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' })],
    });
    renderer.renderSubtitle({ text: 'アクア', tokens: null });

    assert.equal(subtitleRoot.replaceChildrenCalls, 1);
    assert.equal(collectWordNodes(subtitleRoot).length, 1);
    assert.equal(subtitleRoot.textContent, 'アクア');
  } finally {
    restoreDocument();
  }
});

test('renderSubtitle accepts repeated plain payload after style invalidates tokenized render', () => {
  const restoreDocument = installFakeDocument();
  try {
    const subtitleRoot = new FakeElement('div');
    const ctx = {
      state: createRendererState(),
      dom: {
        subtitleRoot,
        subtitleContainer: new FakeElement('div'),
        secondarySubRoot: new FakeElement('div'),
        secondarySubContainer: new FakeElement('div'),
      },
    } as never;

    const renderer = createSubtitleRenderer(ctx);
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' })],
    });
    renderer.applySubtitleStyle({ fontColor: '#fff' } as never);
    renderer.renderSubtitle({ text: 'アクア', tokens: null });

    assert.equal(subtitleRoot.replaceChildrenCalls, 2);
    assert.equal(collectWordNodes(subtitleRoot).length, 0);
    assert.equal(subtitleRoot.textContent, 'アクア');
  } finally {
    restoreDocument();
  }
});

test('renderSubtitle re-renders identical text after style changes affect token output', () => {
  const restoreDocument = installFakeDocument();
  try {
    const subtitleRoot = new FakeElement('div');
    const ctx = {
      state: {
        ...createRendererState(),
        nameMatchEnabled: false,
      },
      dom: {
        subtitleRoot,
        subtitleContainer: new FakeElement('div'),
        secondarySubRoot: new FakeElement('div'),
        secondarySubContainer: new FakeElement('div'),
      },
    } as never;
    const subtitle = {
      text: 'アクア',
      tokens: [
        {
          ...createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' }),
          isNameMatch: true,
        } as MergedToken,
      ],
    };

    const renderer = createSubtitleRenderer(ctx);
    renderer.renderSubtitle(subtitle);
    renderer.applySubtitleStyle({ nameMatchEnabled: true } as never);
    renderer.renderSubtitle(subtitle);

    const [word] = collectWordNodes(subtitleRoot);
    assert.equal(subtitleRoot.replaceChildrenCalls, 2);
    assert.ok(word?.className.includes('word-name-match'));
  } finally {
    restoreDocument();
  }
});

test('renderer content security policy allows data URL character images', () => {
  const htmlPath = path.join(process.cwd(), 'src', 'renderer', 'index.html');
  const htmlText = fs.readFileSync(htmlPath, 'utf-8');
  const cspMatch = htmlText.match(/http-equiv="Content-Security-Policy"[\s\S]*?content="([^"]+)"/);

  assert.ok(cspMatch, 'renderer CSP meta tag should exist');
  assert.match(cspMatch[1] ?? '', /(?:^|;)\s*img-src\s+[^;]*\bdata:/);
});

test('applySubtitleStyle stores secondary background styles in hover-aware css variables', () => {
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
      secondary: {
        backgroundColor: 'rgba(20, 22, 34, 0.78)',
        backdropFilter: 'blur(6px)',
        fontWeight: '600',
      },
    } as never);

    const secondaryStyleValues = (
      secondarySubContainer.style as unknown as {
        values?: Map<string, string>;
        backgroundColor?: string;
        backdropFilter?: string;
      }
    ).values;
    assert.equal(
      secondaryStyleValues?.get('--secondary-sub-background-color'),
      'rgba(20, 22, 34, 0.78)',
    );
    assert.equal(secondaryStyleValues?.get('--secondary-sub-backdrop-filter'), 'blur(6px)');
    assert.equal(
      (secondarySubContainer.style as unknown as { backgroundColor?: string }).backgroundColor,
      undefined,
    );
    assert.equal(
      (secondarySubContainer.style as unknown as { backdropFilter?: string }).backdropFilter,
      undefined,
    );
    assert.equal((secondarySubRoot.style as unknown as { fontWeight?: string }).fontWeight, '600');
  } finally {
    restoreDocument();
  }
});

test('applySubtitleStyle applies primary and secondary css declaration objects', () => {
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
      fontSize: 35,
      css: {
        'font-size': '42px',
        'text-wrap': 'balance',
        '--subtitle-outline': '1px',
      },
      secondary: {
        fontSize: 24,
        css: {
          'font-size': '28px',
          'text-transform': 'uppercase',
        },
      },
    } as never);

    const primaryValues = (subtitleRoot.style as unknown as { values?: Map<string, string> })
      .values;
    const secondaryValues = (secondarySubRoot.style as unknown as { values?: Map<string, string> })
      .values;

    assert.equal(primaryValues?.get('font-size'), '42px');
    assert.equal(primaryValues?.get('text-wrap'), 'balance');
    assert.equal(primaryValues?.get('--subtitle-outline'), '1px');
    assert.equal(secondaryValues?.get('font-size'), '28px');
    assert.equal(secondaryValues?.get('text-transform'), 'uppercase');
  } finally {
    restoreDocument();
  }
});

test('applySubtitleStyle removes css declarations missing from later updates', () => {
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
      css: {
        'font-size': '42px',
        'text-wrap': 'balance',
      },
      secondary: {
        css: {
          'text-transform': 'uppercase',
        },
      },
    } as never);
    renderer.applySubtitleStyle({
      css: {
        'font-size': '44px',
      },
      secondary: {
        css: {},
      },
    } as never);

    const primaryValues = (subtitleRoot.style as unknown as { values?: Map<string, string> })
      .values;
    const secondaryValues = (secondarySubRoot.style as unknown as { values?: Map<string, string> })
      .values;

    assert.equal(primaryValues?.get('font-size'), '44px');
    assert.equal(primaryValues?.has('text-wrap'), false);
    assert.equal(secondaryValues?.has('text-transform'), false);
  } finally {
    restoreDocument();
  }
});

test('annotated subtitle tokens inherit configured base subtitle typography', () => {
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
      fontFamily: 'M PLUS 1 Medium, Source Han Sans JP, Noto Sans CJK JP',
      fontSize: 35,
      fontColor: '#cad3f5',
      fontWeight: 700,
      lineHeight: 1.35,
      letterSpacing: '-0.01em',
      textRendering: 'geometricPrecision',
      textShadow: '3px 0 0 #000, -3px 0 0 #000, 0 3px 0 #000, 0 -3px 0 #000, 2px 2px 0 #000',
      frequencyDictionary: {
        enabled: true,
        topX: 10000,
        mode: 'single',
        singleColor: '#f5a97f',
      },
      enableJlpt: true,
      jlptColors: {
        N1: '#ed8796',
        N2: '#f5a97f',
        N3: '#f9e2af',
        N4: '#a6e3a1',
        N5: '#8aadf4',
      },
      nPlusOneColor: '#c6a0f6',
      knownWordColor: '#a6da95',
    } as never);

    renderer.renderSubtitle({
      text: 'お礼をされるようなことしてない',
      tokens: [
        createToken({ surface: 'お礼', isKnown: true }),
        createToken({ surface: 'を' }),
        createToken({ surface: 'される', jlptLevel: 'N4' }),
        createToken({ surface: 'ような', frequencyRank: 15 }),
      ],
    });

    const rootStyle = subtitleRoot.style as unknown as Record<string, string>;
    assert.equal(rootStyle.fontFamily, 'M PLUS 1 Medium, Source Han Sans JP, Noto Sans CJK JP');
    assert.equal(rootStyle.fontSize, '35px');
    assert.equal(rootStyle.color, '#cad3f5');
    assert.equal(rootStyle.fontWeight, '700');
    assert.equal(rootStyle.lineHeight, '1.35');
    assert.equal(rootStyle.letterSpacing, '-0.01em');
    assert.equal(rootStyle.textRendering, 'geometricPrecision');
    assert.match(rootStyle.textShadow ?? '', /3px 0 0 #000/);

    const wordNodes = collectWordNodes(subtitleRoot);
    assert.deepEqual(
      wordNodes.map((node) => [node.textContent, node.className]),
      [
        ['お礼', 'word word-known'],
        ['を', 'word'],
        ['される', 'word word-jlpt-n4'],
        ['ような', 'word word-frequency-single'],
      ],
    );
    for (const wordNode of wordNodes) {
      const tokenStyle = wordNode.style as unknown as Record<string, string>;
      assert.equal(tokenStyle.fontFamily, undefined);
      assert.equal(tokenStyle.fontSize, undefined);
      assert.equal(tokenStyle.fontWeight, undefined);
      assert.equal(tokenStyle.lineHeight, undefined);
      assert.equal(tokenStyle.letterSpacing, undefined);
      assert.equal(tokenStyle.textRendering, undefined);
      assert.equal(tokenStyle.textShadow, undefined);
    }
  } finally {
    restoreDocument();
  }
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
  const nameToken = createToken({ surface: 'アクア', frequencyRank: 20 }) as MergedToken & {
    isNameMatch?: boolean;
  };
  nameToken.isNameMatch = true;

  assert.equal(getFrequencyRankLabelForToken(frequencyToken, settings), '20');
  assert.equal(getFrequencyRankLabelForToken(knownToken, settings), '20');
  assert.equal(getFrequencyRankLabelForToken(nPlusOneToken, settings), '20');
  assert.equal(getFrequencyRankLabelForToken(outOfRangeToken, settings), null);
  assert.equal(
    getFrequencyRankLabelForToken(nameToken, { ...settings, nameMatchEnabled: true }),
    null,
  );
});

test('getJlptLevelLabelForToken returns level when token has jlpt metadata', () => {
  const jlptToken = createToken({ surface: '語彙', jlptLevel: 'N2' });
  const noJlptToken = createToken({ surface: '語彙' });
  const nameToken = createToken({ surface: 'アクア', jlptLevel: 'N5' }) as MergedToken & {
    isNameMatch?: boolean;
  };
  nameToken.isNameMatch = true;

  assert.equal(getJlptLevelLabelForToken(jlptToken), 'N2');
  assert.equal(getJlptLevelLabelForToken(noJlptToken), null);
  assert.equal(getJlptLevelLabelForToken(nameToken, { nameMatchEnabled: true }), null);
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

test('applySubtitleStyle keeps transparent hover token background', () => {
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
      hoverTokenBackgroundColor: 'transparent',
    } as never);

    const rootStyleValues = (subtitleRoot.style as unknown as { values?: Map<string, string> })
      .values;
    assert.equal(rootStyleValues?.get('--subtitle-hover-token-background-color'), 'transparent');
  } finally {
    restoreDocument();
  }
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

test('renderSubtitle keeps excluded interjection tokens hoverable while rendering them without annotation styling', () => {
  const restoreDocument = installFakeDocument();

  try {
    const subtitleRoot = new FakeElement('div');
    const secondaryRoot = new FakeElement('div');
    const renderer = createSubtitleRenderer({
      dom: {
        subtitleRoot,
        secondarySubtitleRoot: secondaryRoot,
      },
      config: {
        subtitleStyle: {},
        frequencyDictionary: {
          colorTopX: 1000,
          colorMode: 'single',
          colorSingle: '#f5a97f',
          colorBanded: ['#ed8796', '#f5a97f', '#f9e2af', '#8bd5ca', '#8aadf4'],
        },
        secondarySubtitles: { mode: 'hidden' },
      },
      logger: {
        info: () => {},
        warn: () => {},
        error: () => {},
        debug: () => {},
      },
      runtime: {
        secondaryMode: 'hidden' as const,
        shouldToggleMouseIgnore: false,
      },
      state: createRendererState(),
    } as never);

    renderer.renderSubtitle({
      text: 'ぐはっ 猫',
      tokens: [
        createToken({ surface: 'ぐはっ', headword: 'ぐはっ', reading: 'ぐはっ' }),
        createToken({ surface: '猫', headword: '猫', reading: 'ねこ' }),
      ],
    });

    assert.equal(subtitleRoot.textContent, 'ぐはっ 猫');
    assert.deepEqual(
      collectWordNodes(subtitleRoot).map((node) => [node.textContent, node.dataset.tokenIndex]),
      [
        ['ぐはっ', '0'],
        ['猫', '1'],
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

test('normalizeSubtitleForDisplay always breaks between simultaneous cues', () => {
  // The blank line marks two distinct cues on screen at once. Flattening it would run a
  // sign or a second speaker into the line beside it as one sentence.
  const twoCues =
    '\u6b21\u306f\u9b3c\u5b50\u6bcd\u795e\u524d\u3000\u9b3c\u5b50\u6bcd\u795e\u524d\n\n\u611b\u97f3\u3061\u3083\u3093\u3000\u3082\u3046\u5199\u771f\u4e0a\u3052\u3066\u308b';

  assert.equal(
    normalizeSubtitleForDisplay(twoCues, false),
    '\u6b21\u306f\u9b3c\u5b50\u6bcd\u795e\u524d \u9b3c\u5b50\u6bcd\u795e\u524d\n\u611b\u97f3\u3061\u3083\u3093 \u3082\u3046\u5199\u771f\u4e0a\u3052\u3066\u308b',
  );
  assert.equal(normalizeSubtitleForDisplay(twoCues, true), twoCues.replace('\n\n', '\n'));
});

test('normalizeSubtitleForDisplay preserves CRLF boundaries between simultaneous cues', () => {
  assert.equal(normalizeSubtitleForDisplay('a\r\n\r\nb', false), 'a\nb');
});

test('normalizeSubtitleForDisplay still flattens a wrap inside one cue', () => {
  // A typesetter's \\N inside a single utterance is what preserveLineBreaks governs.
  assert.equal(
    normalizeSubtitleForDisplay(
      '\u5e38\u4eba\u304c\u4f7f\u3048\u3070\\N\u305d\u306e\u5727\u5012\u7684\u306a\u529b\u306b',
      false,
    ),
    '\u5e38\u4eba\u304c\u4f7f\u3048\u3070 \u305d\u306e\u5727\u5012\u7684\u306a\u529b\u306b',
  );
});

test('normalizeSubtitle leaves already-decoded text alone', () => {
  // Primary subtitle text is decoded from ASS once, upstream: by mpv for live lines and
  // by the cue parser for prefetched ones. A brace that survives that is literal text.
  assert.equal(normalizeSubtitle('本文{\\pos(1,2)'), '本文{\\pos(1,2)');
  assert.equal(normalizeSubtitle('  余白  ', false), '  余白  ');
});

test('prepareSecondarySubtitleLines drops ASS vector drawing runs', () => {
  assert.deepEqual(
    prepareSecondarySubtitleLines(
      '{\\an5\\pos(730,1042)\\p1\\blur1}m 20 0 b 10 0 0 10 0 20 b 0 31 10 40 20 40 {\\p0}',
    ),
    [],
  );
  assert.deepEqual(prepareSecondarySubtitleLines('{\\p1}m 0 0 l 10 10{\\p0}本文'), ['本文']);
  assert.deepEqual(prepareSecondarySubtitleLines('{\\pos(960,1068)\\bord3}位置指定'), ['位置指定']);
});

test('shouldRenderTokenizedSubtitle enables token rendering when tokens exist', () => {
  assert.equal(shouldRenderTokenizedSubtitle(5), true);
  assert.equal(shouldRenderTokenizedSubtitle(0), false);
});

test('subtitle annotation CSS underlines JLPT tokens without changing token color', () => {
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
    const plainJlptBlock = extractClassBlock(cssText, `#subtitleRoot .word.word-jlpt-n${level}`);
    // JLPT tagging must never recolor the token text — other annotations own
    // text color. JLPT also must not use `text-decoration: underline`,
    // because Chromium repaints text-decoration during ::selection and the
    // underline would adopt the other annotation's color during a Yomitan
    // lookup. The underline is drawn by `border-bottom`, which is unaffected
    // by ::selection and stays locked on the JLPT level color regardless of
    // popup/selection state.
    assert.doesNotMatch(plainJlptBlock, /(?:^|\n)\s*color\s*:/m);
    assert.doesNotMatch(plainJlptBlock, /text-decoration-line:\s*underline;/);
    assert.doesNotMatch(plainJlptBlock, /text-decoration\s*:[^;]*\bunderline\b/i);
    assert.match(
      plainJlptBlock,
      new RegExp(`border-bottom:\\s*3px\\s+solid\\s+var\\(--subtitle-jlpt-n${level}-color,`),
      `JLPT level must paint a permanent 3px border-bottom in the level color`,
    );

    // JLPT tagging must communicate level *only* via the underline; it must
    // never recolor the token text. Other annotations (known, n+1, frequency,
    // name match) are responsible for token text color.
    const jlptOnlyColorBlock = extractClassBlock(cssText, buildJlptColorSelector(level));
    assert.equal(
      jlptOnlyColorBlock,
      '',
      `word-jlpt-n${level} (without other annotations) must not set text color — JLPT only paints the underline`,
    );
  }

  for (const selector of [
    '#subtitleRoot .word.word-known',
    '#subtitleRoot .word.word-n-plus-one',
    '#subtitleRoot .word.word-name-match',
  ]) {
    const block = extractClassBlock(cssText, selector);
    assert.match(block, /color:\s*var\(/);
    assert.doesNotMatch(block, /text-shadow\s*:/);
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
  assert.match(subtitleRootBlock, /--subtitle-hover-token-background-color:\s*transparent;/);
  assert.match(subtitleRootBlock, /-webkit-text-fill-color:\s*currentColor;/);

  const charBlock = extractClassBlock(cssText, '#subtitleRoot .c');
  assert.match(charBlock, /-webkit-text-fill-color:\s*currentColor\s*!important;/);

  const wordBlock = extractClassBlock(cssText, '#subtitleRoot .word');
  assert.match(wordBlock, /-webkit-text-fill-color:\s*currentColor\s*!important;/);

  const characterImageTokenBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word.word-character-image-token',
  );
  assert.match(characterImageTokenBlock, /display:\s*inline-block;/);
  assert.match(characterImageTokenBlock, /position:\s*relative;/);
  assert.match(characterImageTokenBlock, /padding-left:\s*1\.08em;/);
  assert.match(characterImageTokenBlock, /margin-left:\s*0\.18em;/);

  const characterImageBlock = extractClassBlock(cssText, '#subtitleRoot .word-character-image');
  assert.match(characterImageBlock, /position:\s*absolute;/);
  assert.match(characterImageBlock, /top:\s*50%;/);
  assert.match(characterImageBlock, /transform:\s*translateY\(calc\(-50%\s*\+\s*0\.05em\)\);/);

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

  const plainWordHoverBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word:not(.word-known):not(.word-n-plus-one):not(.word-name-match):not(.word-frequency-single):not(.word-frequency-band-1):not(.word-frequency-band-2):not(.word-frequency-band-3):not(.word-frequency-band-4):not(.word-frequency-band-5):hover',
  );
  assert.match(
    plainWordHoverBlock,
    /background:\s*var\(--subtitle-hover-token-background-color,\s*transparent\);/,
  );
  assert.match(
    plainWordHoverBlock,
    /color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );
  assert.match(
    plainWordHoverBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  const coloredWordHoverBlock = extractClassBlock(cssText, '#subtitleRoot .word.word-known:hover');
  assert.match(
    coloredWordHoverBlock,
    /background:\s*var\(--subtitle-hover-token-background-color,\s*transparent\);/,
  );
  assert.match(coloredWordHoverBlock, /border-radius:\s*3px;/);
  assert.match(coloredWordHoverBlock, /filter:\s*brightness\(1\.18\) saturate\(1\.08\);/);
  assert.doesNotMatch(coloredWordHoverBlock, /font-weight\s*:/);
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

  const jlptOnlyHoverBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word:is(.word-jlpt-n1, .word-jlpt-n2, .word-jlpt-n3, .word-jlpt-n4, .word-jlpt-n5):not(.word-known):not(.word-n-plus-one):not(.word-name-match):not(.word-frequency-single):not(.word-frequency-band-1):not(.word-frequency-band-2):not(.word-frequency-band-3):not(.word-frequency-band-4):not(.word-frequency-band-5):hover',
  );
  assert.match(
    jlptOnlyHoverBlock,
    /color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );
  assert.match(
    jlptOnlyHoverBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  const jlptOnlySelectionBlock = extractClassBlock(
    cssText,
    '#subtitleRoot .word:is(.word-jlpt-n1, .word-jlpt-n2, .word-jlpt-n3, .word-jlpt-n4, .word-jlpt-n5):not(.word-known):not(.word-n-plus-one):not(.word-name-match):not(.word-frequency-single):not(.word-frequency-band-1):not(.word-frequency-band-2):not(.word-frequency-band-3):not(.word-frequency-band-4):not(.word-frequency-band-5)::selection',
  );
  assert.match(
    jlptOnlySelectionBlock,
    /color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );
  assert.match(
    jlptOnlySelectionBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  for (let level = 1; level <= 5; level += 1) {
    const jlptSelectionLockBlock = extractClassBlock(
      cssText,
      `#subtitleRoot .word.word-jlpt-n${level}::selection`,
    );
    assert.ok(jlptSelectionLockBlock.length > 0, `word-jlpt-n${level} selection lock should exist`);
    assert.match(
      jlptSelectionLockBlock,
      new RegExp(
        `text-decoration-color:\\s*var\\(--subtitle-jlpt-n${level}-color,\\s*#[0-9a-f]{6}\\)\\s*!important;`,
        'i',
      ),
    );

    for (const annotationClass of [
      'word-known',
      'word-n-plus-one',
      'word-name-match',
      'word-frequency-single',
      'word-frequency-band-2',
    ]) {
      const combinedAnnotationBlock = extractClassBlock(
        cssText,
        `#subtitleRoot .word.word-jlpt-n${level}.${annotationClass}`,
      );
      assert.match(
        combinedAnnotationBlock,
        new RegExp(
          `text-decoration-color:\\s*var\\(--subtitle-jlpt-n${level}-color,\\s*#[0-9a-f]{6}\\)\\s*!important;`,
          'i',
        ),
        `combined JLPT ${annotationClass} selector should lock underline color`,
      );
    }

    const jlptCharHoverBlock = extractClassBlock(
      cssText,
      `#subtitleRoot .word.word-jlpt-n${level} .c:hover`,
    );
    assert.match(
      jlptCharHoverBlock,
      new RegExp(
        `text-decoration-color:\\s*var\\(--subtitle-jlpt-n${level}-color,\\s*#[0-9a-f]{6}\\)\\s*!important;`,
        'i',
      ),
      'JLPT character hover selector should lock underline color',
    );
  }

  const selectionBlock = extractClassBlock(cssText, '#subtitleRoot::selection');
  assert.match(
    selectionBlock,
    /background:\s*var\(--subtitle-hover-token-background-color,\s*transparent\);/,
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
    /background:\s*var\(--subtitle-hover-token-background-color,\s*transparent\)\s*!important;/,
  );
  assert.match(
    descendantSelectionBlock,
    /color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );
  assert.match(
    descendantSelectionBlock,
    /-webkit-text-fill-color:\s*var\(--subtitle-hover-token-color,\s*#f4dbd6\)\s*!important;/,
  );

  const secondaryContainerBlock = extractClassBlock(cssText, '#secondarySubContainer');
  assert.match(
    secondaryContainerBlock,
    /background:\s*var\(--secondary-sub-background-color,\s*transparent\);/,
  );
  assert.match(
    secondaryContainerBlock,
    /backdrop-filter:\s*var\(--secondary-sub-backdrop-filter,\s*none\);/,
  );

  const secondaryRootBlock = extractClassBlock(cssText, '#secondarySubRoot');
  assert.match(secondaryRootBlock, /-webkit-text-stroke:\s*0\.45px rgba\(0,\s*0,\s*0,\s*0\.7\);/);
  assert.match(
    secondaryRootBlock,
    /text-shadow:\s*-1px -1px 2px rgba\(0,\s*0,\s*0,\s*0\.95\),\s*1px -1px 2px rgba\(0,\s*0,\s*0,\s*0\.95\),\s*-1px 1px 2px rgba\(0,\s*0,\s*0,\s*0\.95\),\s*1px 1px 2px rgba\(0,\s*0,\s*0,\s*0\.95\),\s*0 0 8px rgba\(0,\s*0,\s*0,\s*0\.5\);/,
  );

  const secondaryHoverBaseBlock = extractClassBlock(
    cssText,
    '#secondarySubContainer.secondary-sub-hover #secondarySubRoot',
  );
  assert.match(secondaryHoverBaseBlock, /background:\s*transparent;/);

  const primaryHoverBlock = extractClassBlock(cssText, '#subtitleContainer.primary-sub-hover');
  assert.match(primaryHoverBlock, /opacity:\s*0;/);
  assert.match(primaryHoverBlock, /pointer-events:\s*auto;/);

  const primaryHoverVisibleBlock = extractClassBlock(
    cssText,
    '#subtitleContainer.primary-sub-hover:hover',
  );
  assert.match(primaryHoverVisibleBlock, /opacity:\s*1;/);

  const primaryHoverYomitanPopupVisibleBlock = extractClassBlock(
    cssText,
    'body.primary-sub-visible-on-yomitan-popup #subtitleContainer.primary-sub-hover',
  );
  assert.match(primaryHoverYomitanPopupVisibleBlock, /opacity:\s*1;/);

  const secondaryEmbeddedHoverBlock = extractClassBlock(
    cssText,
    'body.subtitle-sidebar-embedded-open #secondarySubContainer.secondary-sub-hover',
  );
  assert.match(secondaryEmbeddedHoverBlock, /right:\s*var\(--subtitle-sidebar-reserved-width\);/);
  assert.match(secondaryEmbeddedHoverBlock, /max-width:\s*none;/);
  assert.match(secondaryEmbeddedHoverBlock, /transform:\s*none;/);
  assert.doesNotMatch(
    secondaryEmbeddedHoverBlock,
    /transform:\s*translateX\(calc\(var\(--subtitle-sidebar-reserved-width\)\s*\*\s*-0\.5\)\);/,
  );

  const secondaryHoverWindowsBlock = extractClassBlock(
    cssText,
    'body.platform-windows #secondarySubContainer.secondary-sub-hover',
  );
  assert.match(secondaryHoverWindowsBlock, /top:\s*40px;/);
  assert.match(secondaryHoverWindowsBlock, /padding-top:\s*0;/);

  const sidebarSettingsModalBlock = extractClassBlock(
    cssText,
    'body.settings-modal-open .subtitle-sidebar-modal',
  );
  assert.match(sidebarSettingsModalBlock, /display:\s*none !important;/);
  assert.match(sidebarSettingsModalBlock, /pointer-events:\s*none !important;/);

  const subtitleSidebarListBlock = extractClassBlock(cssText, '.subtitle-sidebar-list');
  assert.doesNotMatch(subtitleSidebarListBlock, /scroll-behavior:\s*smooth;/);

  const secondaryHoverVisibleBlock = extractClassBlock(
    cssText,
    '#secondarySubContainer.secondary-sub-hover:hover #secondarySubRoot',
  );
  assert.match(
    secondaryHoverVisibleBlock,
    /background:\s*var\(--secondary-sub-background-color,\s*transparent\);/,
  );
  assert.match(
    secondaryHoverVisibleBlock,
    /backdrop-filter:\s*var\(--secondary-sub-backdrop-filter,\s*none\);/,
  );

  const secondaryHoverActiveBlock = extractClassBlock(
    cssText,
    '#secondarySubContainer.secondary-sub-hover.secondary-sub-hover-active',
  );
  assert.match(secondaryHoverActiveBlock, /opacity:\s*1;/);

  const secondaryHoverActiveRootBlock = extractClassBlock(
    cssText,
    '#secondarySubContainer.secondary-sub-hover.secondary-sub-hover-active #secondarySubRoot',
  );
  assert.match(
    secondaryHoverActiveRootBlock,
    /background:\s*var\(--secondary-sub-background-color,\s*transparent\);/,
  );
  assert.match(
    secondaryHoverActiveRootBlock,
    /backdrop-filter:\s*var\(--secondary-sub-backdrop-filter,\s*none\);/,
  );

  assert.doesNotMatch(
    cssText,
    /body\.layer-visible\s+#secondarySubContainer\s*\{[^}]*display:\s*none/i,
  );
});

test('prepareSecondarySubtitleLines collapses exact short copies in stacks', () => {
  assert.deepEqual(prepareSecondarySubtitleLines('Your\\NYour\\NYour\\NYour\\Nmosaic'), [
    'Your',
    'mosaic',
  ]);
  assert.deepEqual(prepareSecondarySubtitleLines('One line\\NAnother line'), [
    'One line',
    'Another line',
  ]);
});

test('prepareSecondarySubtitleLines collapses exact short sign copies beside dialogue', () => {
  const liveText = "And for today's sports festival...\nEntrance\nEntrance";

  assert.deepEqual(prepareSecondarySubtitleLines(liveText), [
    "And for today's sports festival...",
    'Entrance',
  ]);
});

test('prepareSecondarySubtitleLines collapses karaoke syllable spam into one deduped line', () => {
  // Karaoke-typeset OP/ED: one ASS event per syllable, duplicated across layers,
  // joined with \N by mpv's secondary-sub-text.
  const karaoke = ['ya', 'This', 'ya', 'This', 'ya', 'This', 'no', 'ma', 'ups', 'ma', 'ups'].join(
    '\\N',
  );

  assert.deepEqual(prepareSecondarySubtitleLines(karaoke), ['ya This no ma ups']);
});

test('prepareSecondarySubtitleLines collapses exact repeated short lines', () => {
  const dialogue = ['Wait', 'Wait', 'Wait'];

  assert.deepEqual(prepareSecondarySubtitleLines(dialogue.join('\\N')), ['Wait']);
});

test('prepareSecondarySubtitleLines collapses punctuation variants of a full-sentence fallback', () => {
  const dialogue = 'A question veiled as an insult!';
  const positionedSign = 'A question veiled as an insult';

  assert.deepEqual(prepareSecondarySubtitleLines([dialogue, positionedSign].join('\\N')), [
    dialogue,
  ]);
});

test('prepareSecondarySubtitleLines preserves short simultaneous dialogue without repeats', () => {
  const dialogue = ['Wait', 'Go!', 'No!', 'Run!'];

  assert.deepEqual(prepareSecondarySubtitleLines(dialogue.join('\\N')), dialogue);
});

test('prepareSecondarySubtitleLines preserves distinct short lines with internal whitespace', () => {
  assert.deepEqual(prepareSecondarySubtitleLines('AB\\NA B'), ['AB', 'A B']);
});

test('prepareSecondarySubtitleLines keeps normal dialogue lines intact', () => {
  const dialogue = ' I never expected this. \\N\\N But here we are. ';

  assert.deepEqual(prepareSecondarySubtitleLines(dialogue), [
    'I never expected this.',
    'But here we are.',
  ]);
});

test('prepareSecondarySubtitleLines does not collapse many long simultaneous lines', () => {
  const lines = Array.from({ length: 9 }, (_, i) => `This is a full sentence number ${i}.`);

  assert.deepEqual(prepareSecondarySubtitleLines(lines.join('\\N')), lines);
});

test('prepareSecondarySubtitleLines strips ASS override tags and handles empty input', () => {
  assert.deepEqual(prepareSecondarySubtitleLines('{\\an8}Sign text'), ['Sign text']);
  assert.deepEqual(prepareSecondarySubtitleLines(''), []);
  assert.deepEqual(prepareSecondarySubtitleLines('{\\an8}'), []);
});

test('secondary subtitle root CSS does not clip long subtitle stacks', () => {
  const srcCssPath = path.join(process.cwd(), 'src', 'renderer', 'style.css');
  const cssText = fs.readFileSync(srcCssPath, 'utf-8');

  const secondaryRootBlock = extractClassBlock(cssText, '#secondarySubRoot');
  assert.doesNotMatch(secondaryRootBlock, /max-height\s*:/);
  assert.doesNotMatch(secondaryRootBlock, /overflow\s*:\s*hidden/);
});

test('applySubtitleStyle sets known-word maturity color variables', () => {
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
      knownWordMaturityColors: {
        new: '#111111',
        learning: '#222222',
        young: '#333333',
        mature: '#444444',
      },
    } as never);

    const values = (subtitleRoot.style as unknown as { values?: Map<string, string> }).values;
    assert.equal(values?.get('--subtitle-maturity-new-color'), '#111111');
    assert.equal(values?.get('--subtitle-maturity-learning-color'), '#222222');
    assert.equal(values?.get('--subtitle-maturity-young-color'), '#333333');
    assert.equal(values?.get('--subtitle-maturity-mature-color'), '#444444');
  } finally {
    restoreDocument();
  }
});

test('applySubtitleStyle falls back to default maturity colors', () => {
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
    renderer.applySubtitleStyle({} as never);

    const values = (subtitleRoot.style as unknown as { values?: Map<string, string> }).values;
    assert.equal(values?.get('--subtitle-maturity-new-color'), '#ee99a0');
    assert.equal(values?.get('--subtitle-maturity-learning-color'), '#b7bdf8');
    assert.equal(values?.get('--subtitle-maturity-young-color'), '#91d7e3');
    assert.equal(values?.get('--subtitle-maturity-mature-color'), '#a6da95');
  } finally {
    restoreDocument();
  }
});
