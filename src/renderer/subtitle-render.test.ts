import test from 'node:test';
import assert from 'node:assert/strict';

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
} from './subtitle-render.js';
import { createToken } from './subtitle-render-test-helpers.js';
import { createRendererState, type RendererState } from './state.js';

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
  // Plain assignments (style.fontSize = ...) land as own properties.
  [property: string]: unknown;

  // Values written through setProperty (custom properties and css declaration objects).
  readonly values = new Map<string, string>();

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
  src?: string;
  alt?: string;
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

    return this.childNodes.map((child) => child.textContent).join('');
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

type FakeDom = {
  subtitleRoot: FakeElement;
  subtitleContainer: FakeElement;
  secondarySubRoot: FakeElement;
  secondarySubContainer: FakeElement;
};

// Installs a fake `document`, builds a subtitle renderer over fresh fake roots, runs `fn`,
// then restores the previous `document`.
function withRenderer(
  stateOverrides: Partial<RendererState>,
  fn: (env: { renderer: ReturnType<typeof createSubtitleRenderer>; dom: FakeDom }) => void,
): void {
  const previousDocument = (globalThis as { document?: unknown }).document;
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createDocumentFragment: () => new FakeDocumentFragment(),
      createElement: (tagName: string) => new FakeElement(tagName),
      createTextNode: (text: string) => new FakeTextNode(text),
    },
  });

  try {
    const dom: FakeDom = {
      subtitleRoot: new FakeElement('div'),
      subtitleContainer: new FakeElement('div'),
      secondarySubRoot: new FakeElement('div'),
      secondarySubContainer: new FakeElement('div'),
    };
    const ctx = { state: { ...createRendererState(), ...stateOverrides }, dom };
    fn({ renderer: createSubtitleRenderer(ctx as never), dom });
  } finally {
    Object.defineProperty(globalThis, 'document', {
      configurable: true,
      value: previousDocument,
    });
  }
}

function collectWordNodes(root: FakeElement): FakeElement[] {
  return root.childNodes.filter(
    (child): child is FakeElement =>
      child instanceof FakeElement && child.className.includes('word'),
  );
}

const namedToken = (extra: Partial<MergedToken> = {}): MergedToken => ({
  ...createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' }),
  isNameMatch: true,
  ...extra,
});

test('renderSubtitle injects circular character image for annotated name matches', () => {
  withRenderer({ nameMatchEnabled: true }, ({ renderer, dom }) => {
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [
        namedToken({ characterImage: { src: 'data:image/png;base64,AAAA', alt: 'アクア' } }),
      ],
    });

    const [word] = collectWordNodes(dom.subtitleRoot);
    assert.ok(word);
    assert.equal(word.className, 'word word-name-match word-character-image-token');
    assert.equal(word.textContent, 'アクア');
    const image = word.childNodes[0];
    assert.ok(image instanceof FakeElement);
    assert.equal(image.tagName, 'img');
    assert.equal(image.className, 'word-character-image');
    assert.equal(image.src, 'data:image/png;base64,AAAA');
    assert.equal(image.alt, 'アクア');
  });
});

test('renderSubtitle skips character image when name-match rendering is disabled', () => {
  withRenderer({ nameMatchEnabled: false }, ({ renderer, dom }) => {
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [
        namedToken({ characterImage: { src: 'data:image/png;base64,AAAA', alt: 'アクア' } }),
      ],
    });

    const [word] = collectWordNodes(dom.subtitleRoot);
    assert.ok(word);
    assert.equal(word.className, 'word');
    assert.equal(word.textContent, 'アクア');
    assert.equal(word.childNodes.length, 0);
  });
});

test('renderSubtitle skips identical primary subtitle DOM replacement', () => {
  withRenderer({}, ({ renderer, dom }) => {
    renderer.renderSubtitle({ text: '字幕', tokens: null });
    renderer.renderSubtitle({ text: '字幕', tokens: null });
    renderer.renderSubtitle({ text: '字幕2', tokens: null });

    assert.equal(dom.subtitleRoot.replaceChildrenCalls, 2);
    assert.equal(dom.subtitleRoot.textContent, '字幕2');
  });
});

test('renderSubtitle keeps tokenized subtitle when stale plain payload repeats same text', () => {
  withRenderer({}, ({ renderer, dom }) => {
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' })],
    });
    renderer.renderSubtitle({ text: 'アクア', tokens: null });

    assert.equal(dom.subtitleRoot.replaceChildrenCalls, 1);
    assert.equal(collectWordNodes(dom.subtitleRoot).length, 1);
    assert.equal(dom.subtitleRoot.textContent, 'アクア');
  });
});

test('renderSubtitle accepts repeated plain payload after style invalidates tokenized render', () => {
  withRenderer({}, ({ renderer, dom }) => {
    renderer.renderSubtitle({
      text: 'アクア',
      tokens: [createToken({ surface: 'アクア', headword: 'アクア', reading: 'あくあ' })],
    });
    renderer.applySubtitleStyle({ fontColor: '#fff' } as never);
    renderer.renderSubtitle({ text: 'アクア', tokens: null });

    assert.equal(dom.subtitleRoot.replaceChildrenCalls, 2);
    assert.equal(collectWordNodes(dom.subtitleRoot).length, 0);
    assert.equal(dom.subtitleRoot.textContent, 'アクア');
  });
});

test('renderSubtitle re-renders identical text after style changes affect token output', () => {
  withRenderer({ nameMatchEnabled: false }, ({ renderer, dom }) => {
    const subtitle = { text: 'アクア', tokens: [namedToken()] };

    renderer.renderSubtitle(subtitle);
    renderer.applySubtitleStyle({ nameMatchEnabled: true } as never);
    renderer.renderSubtitle(subtitle);

    const [word] = collectWordNodes(dom.subtitleRoot);
    assert.equal(dom.subtitleRoot.replaceChildrenCalls, 2);
    assert.ok(word?.className.includes('word-name-match'));
  });
});

test('renderSubtitle preserves unsupported punctuation while keeping it non-interactive', () => {
  withRenderer({}, ({ renderer, dom }) => {
    renderer.renderSubtitle({
      text: 'えっ！？マジ',
      tokens: [createToken({ surface: 'えっ' }), createToken({ surface: 'マジ' })],
    });

    assert.equal(dom.subtitleRoot.textContent, 'えっ！？マジ');
    assert.deepEqual(
      collectWordNodes(dom.subtitleRoot).map((node) => [node.textContent, node.dataset.tokenIndex]),
      [
        ['えっ', '0'],
        ['マジ', '1'],
      ],
    );
  });
});

test('applySubtitleStyle stores secondary background styles in hover-aware css variables', () => {
  withRenderer({}, ({ renderer, dom }) => {
    renderer.applySubtitleStyle({
      secondary: {
        backgroundColor: 'rgba(20, 22, 34, 0.78)',
        backdropFilter: 'blur(6px)',
        fontWeight: '600',
      },
    } as never);

    const containerStyle = dom.secondarySubContainer.style;
    assert.equal(
      containerStyle.values.get('--secondary-sub-background-color'),
      'rgba(20, 22, 34, 0.78)',
    );
    assert.equal(containerStyle.values.get('--secondary-sub-backdrop-filter'), 'blur(6px)');
    assert.equal(containerStyle.backgroundColor, undefined);
    assert.equal(containerStyle.backdropFilter, undefined);
    assert.equal(dom.secondarySubRoot.style.fontWeight, '600');
  });
});

test('applySubtitleStyle applies primary and secondary css declaration objects', () => {
  withRenderer({}, ({ renderer, dom }) => {
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

    const primaryValues = dom.subtitleRoot.style.values;
    const secondaryValues = dom.secondarySubRoot.style.values;
    assert.equal(primaryValues.get('font-size'), '42px');
    assert.equal(primaryValues.get('text-wrap'), 'balance');
    assert.equal(primaryValues.get('--subtitle-outline'), '1px');
    assert.equal(secondaryValues.get('font-size'), '28px');
    assert.equal(secondaryValues.get('text-transform'), 'uppercase');
  });
});

test('applySubtitleStyle removes css declarations missing from later updates', () => {
  withRenderer({}, ({ renderer, dom }) => {
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

    const primaryValues = dom.subtitleRoot.style.values;
    assert.equal(primaryValues.get('font-size'), '44px');
    assert.equal(primaryValues.has('text-wrap'), false);
    assert.equal(dom.secondarySubRoot.style.values.has('text-transform'), false);
  });
});

test('annotated subtitle tokens inherit configured base subtitle typography', () => {
  withRenderer({}, ({ renderer, dom }) => {
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

    const rootStyle = dom.subtitleRoot.style;
    assert.equal(rootStyle.fontFamily, 'M PLUS 1 Medium, Source Han Sans JP, Noto Sans CJK JP');
    assert.equal(rootStyle.fontSize, '35px');
    assert.equal(rootStyle.color, '#cad3f5');
    assert.equal(rootStyle.fontWeight, '700');
    assert.equal(rootStyle.lineHeight, '1.35');
    assert.equal(rootStyle.letterSpacing, '-0.01em');
    assert.equal(rootStyle.textRendering, 'geometricPrecision');
    assert.match(String(rootStyle.textShadow), /3px 0 0 #000/);

    const wordNodes = collectWordNodes(dom.subtitleRoot);
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
      const tokenStyle = wordNode.style;
      assert.equal(tokenStyle.fontFamily, undefined);
      assert.equal(tokenStyle.fontSize, undefined);
      assert.equal(tokenStyle.fontWeight, undefined);
      assert.equal(tokenStyle.lineHeight, undefined);
      assert.equal(tokenStyle.letterSpacing, undefined);
      assert.equal(tokenStyle.textRendering, undefined);
      assert.equal(tokenStyle.textShadow, undefined);
    }
  });
});

test('applySubtitleStyle keeps transparent hover token background', () => {
  withRenderer({}, ({ renderer, dom }) => {
    renderer.applySubtitleStyle({ hoverTokenBackgroundColor: 'transparent' } as never);

    assert.equal(
      dom.subtitleRoot.style.values.get('--subtitle-hover-token-background-color'),
      'transparent',
    );
  });
});

test('applySubtitleStyle sets known-word maturity color variables', () => {
  withRenderer({}, ({ renderer, dom }) => {
    renderer.applySubtitleStyle({
      knownWordMaturityColors: {
        new: '#111111',
        learning: '#222222',
        young: '#333333',
        mature: '#444444',
      },
    } as never);

    const values = dom.subtitleRoot.style.values;
    assert.equal(values.get('--subtitle-maturity-new-color'), '#111111');
    assert.equal(values.get('--subtitle-maturity-learning-color'), '#222222');
    assert.equal(values.get('--subtitle-maturity-young-color'), '#333333');
    assert.equal(values.get('--subtitle-maturity-mature-color'), '#444444');
  });
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
