import { ResolvedConfig } from '../../types/config';
import { ResolveContext } from './context';
import {
  asBoolean,
  asColor,
  asCssColor,
  asFrequencyBandedColors,
  asNumber,
  asString,
  isObject,
} from './shared';

export function applySubtitleDomainConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

  if (isObject(src.jimaku)) {
    const apiKey = asString(src.jimaku.apiKey);
    if (apiKey !== undefined) resolved.jimaku.apiKey = apiKey;
    const apiKeyCommand = asString(src.jimaku.apiKeyCommand);
    if (apiKeyCommand !== undefined) resolved.jimaku.apiKeyCommand = apiKeyCommand;
    const apiBaseUrl = asString(src.jimaku.apiBaseUrl);
    if (apiBaseUrl !== undefined) resolved.jimaku.apiBaseUrl = apiBaseUrl;

    const lang = src.jimaku.languagePreference;
    if (lang === 'ja' || lang === 'en' || lang === 'none') {
      resolved.jimaku.languagePreference = lang;
    } else if (lang !== undefined) {
      warn(
        'jimaku.languagePreference',
        lang,
        resolved.jimaku.languagePreference,
        'Expected ja, en, or none.',
      );
    }

    const maxEntryResults = asNumber(src.jimaku.maxEntryResults);
    if (maxEntryResults !== undefined && maxEntryResults > 0) {
      resolved.jimaku.maxEntryResults = Math.floor(maxEntryResults);
    } else if (src.jimaku.maxEntryResults !== undefined) {
      warn(
        'jimaku.maxEntryResults',
        src.jimaku.maxEntryResults,
        resolved.jimaku.maxEntryResults,
        'Expected positive number.',
      );
    }
  }

  if (isObject(src.youtubeSubgen)) {
    const whisperBin = asString(src.youtubeSubgen.whisperBin);
    if (whisperBin !== undefined) {
      resolved.youtubeSubgen.whisperBin = whisperBin;
    } else if (src.youtubeSubgen.whisperBin !== undefined) {
      warn(
        'youtubeSubgen.whisperBin',
        src.youtubeSubgen.whisperBin,
        resolved.youtubeSubgen.whisperBin,
        'Expected string.',
      );
    }

    const whisperModel = asString(src.youtubeSubgen.whisperModel);
    if (whisperModel !== undefined) {
      resolved.youtubeSubgen.whisperModel = whisperModel;
    } else if (src.youtubeSubgen.whisperModel !== undefined) {
      warn(
        'youtubeSubgen.whisperModel',
        src.youtubeSubgen.whisperModel,
        resolved.youtubeSubgen.whisperModel,
        'Expected string.',
      );
    }

    const whisperVadModel = asString(src.youtubeSubgen.whisperVadModel);
    if (whisperVadModel !== undefined) {
      resolved.youtubeSubgen.whisperVadModel = whisperVadModel;
    } else if (src.youtubeSubgen.whisperVadModel !== undefined) {
      warn(
        'youtubeSubgen.whisperVadModel',
        src.youtubeSubgen.whisperVadModel,
        resolved.youtubeSubgen.whisperVadModel,
        'Expected string.',
      );
    }

    const whisperThreads = asNumber(src.youtubeSubgen.whisperThreads);
    if (whisperThreads !== undefined && Number.isInteger(whisperThreads) && whisperThreads > 0) {
      resolved.youtubeSubgen.whisperThreads = whisperThreads;
    } else if (src.youtubeSubgen.whisperThreads !== undefined) {
      warn(
        'youtubeSubgen.whisperThreads',
        src.youtubeSubgen.whisperThreads,
        resolved.youtubeSubgen.whisperThreads,
        'Expected positive integer.',
      );
    }

    const fixWithAi = asBoolean(src.youtubeSubgen.fixWithAi);
    if (fixWithAi !== undefined) {
      resolved.youtubeSubgen.fixWithAi = fixWithAi;
    } else if (src.youtubeSubgen.fixWithAi !== undefined) {
      warn(
        'youtubeSubgen.fixWithAi',
        src.youtubeSubgen.fixWithAi,
        resolved.youtubeSubgen.fixWithAi,
        'Expected boolean.',
      );
    }

    if (isObject(src.youtubeSubgen.ai)) {
      const aiModel = asString(src.youtubeSubgen.ai.model);
      if (aiModel !== undefined) {
        resolved.youtubeSubgen.ai.model = aiModel;
      } else if (src.youtubeSubgen.ai.model !== undefined) {
        warn(
          'youtubeSubgen.ai.model',
          src.youtubeSubgen.ai.model,
          resolved.youtubeSubgen.ai.model,
          'Expected string.',
        );
      }

      const aiSystemPrompt = asString(src.youtubeSubgen.ai.systemPrompt);
      if (aiSystemPrompt !== undefined) {
        resolved.youtubeSubgen.ai.systemPrompt = aiSystemPrompt;
      } else if (src.youtubeSubgen.ai.systemPrompt !== undefined) {
        warn(
          'youtubeSubgen.ai.systemPrompt',
          src.youtubeSubgen.ai.systemPrompt,
          resolved.youtubeSubgen.ai.systemPrompt,
          'Expected string.',
        );
      }
    } else if (src.youtubeSubgen.ai !== undefined) {
      warn('youtubeSubgen.ai', src.youtubeSubgen.ai, resolved.youtubeSubgen.ai, 'Expected object.');
    }

    if (src.youtubeSubgen.primarySubLanguages !== undefined) {
      warn(
        'youtubeSubgen.primarySubLanguages',
        src.youtubeSubgen.primarySubLanguages,
        undefined,
        'Removed. Use youtube.primarySubLanguages instead.',
      );
    }
  }

  if (isObject(src.subtitleStyle)) {
    const fallbackSubtitleStyleEnableJlpt = resolved.subtitleStyle.enableJlpt;
    const fallbackSubtitleStylePrimaryDefaultMode = resolved.subtitleStyle.primaryDefaultMode;
    const fallbackSubtitleStylePreserveLineBreaks = resolved.subtitleStyle.preserveLineBreaks;
    const fallbackSubtitleStyleAutoPauseVideoOnHover = resolved.subtitleStyle.autoPauseVideoOnHover;
    const fallbackSubtitleStyleAutoPauseVideoOnYomitanPopup =
      resolved.subtitleStyle.autoPauseVideoOnYomitanPopup;
    const fallbackSubtitleStyleHoverTokenColor = resolved.subtitleStyle.hoverTokenColor;
    const fallbackSubtitleStyleHoverTokenBackgroundColor =
      resolved.subtitleStyle.hoverTokenBackgroundColor;
    const fallbackSubtitleStyleNameMatchEnabled = resolved.subtitleStyle.nameMatchEnabled;
    const fallbackSubtitleStyleNameMatchColor = resolved.subtitleStyle.nameMatchColor;
    const fallbackSubtitleStyleKnownWordColor = resolved.subtitleStyle.knownWordColor;
    const fallbackSubtitleStyleNPlusOneColor = resolved.subtitleStyle.nPlusOneColor;
    const fallbackFrequencyDictionary = {
      ...resolved.subtitleStyle.frequencyDictionary,
    };
    resolved.subtitleStyle = {
      ...resolved.subtitleStyle,
      ...(src.subtitleStyle as ResolvedConfig['subtitleStyle']),
      frequencyDictionary: {
        ...resolved.subtitleStyle.frequencyDictionary,
        ...(isObject((src.subtitleStyle as { frequencyDictionary?: unknown }).frequencyDictionary)
          ? ((src.subtitleStyle as { frequencyDictionary?: unknown })
              .frequencyDictionary as ResolvedConfig['subtitleStyle']['frequencyDictionary'])
          : {}),
      },
      secondary: {
        ...resolved.subtitleStyle.secondary,
        ...(isObject(src.subtitleStyle.secondary)
          ? (src.subtitleStyle.secondary as ResolvedConfig['subtitleStyle']['secondary'])
          : {}),
      },
    };

    const enableJlpt = asBoolean((src.subtitleStyle as { enableJlpt?: unknown }).enableJlpt);
    if (enableJlpt !== undefined) {
      resolved.subtitleStyle.enableJlpt = enableJlpt;
    } else if ((src.subtitleStyle as { enableJlpt?: unknown }).enableJlpt !== undefined) {
      resolved.subtitleStyle.enableJlpt = fallbackSubtitleStyleEnableJlpt;
      warn(
        'subtitleStyle.enableJlpt',
        (src.subtitleStyle as { enableJlpt?: unknown }).enableJlpt,
        resolved.subtitleStyle.enableJlpt,
        'Expected boolean.',
      );
    }

    const primaryDefaultMode = (src.subtitleStyle as { primaryDefaultMode?: unknown })
      .primaryDefaultMode;
    if (
      primaryDefaultMode === 'hidden' ||
      primaryDefaultMode === 'visible' ||
      primaryDefaultMode === 'hover'
    ) {
      resolved.subtitleStyle.primaryDefaultMode = primaryDefaultMode;
    } else if (primaryDefaultMode !== undefined) {
      resolved.subtitleStyle.primaryDefaultMode = fallbackSubtitleStylePrimaryDefaultMode;
      warn(
        'subtitleStyle.primaryDefaultMode',
        primaryDefaultMode,
        resolved.subtitleStyle.primaryDefaultMode,
        'Expected hidden, visible, or hover.',
      );
    }

    const preserveLineBreaks = asBoolean(
      (src.subtitleStyle as { preserveLineBreaks?: unknown }).preserveLineBreaks,
    );
    if (preserveLineBreaks !== undefined) {
      resolved.subtitleStyle.preserveLineBreaks = preserveLineBreaks;
    } else if (
      (src.subtitleStyle as { preserveLineBreaks?: unknown }).preserveLineBreaks !== undefined
    ) {
      resolved.subtitleStyle.preserveLineBreaks = fallbackSubtitleStylePreserveLineBreaks;
      warn(
        'subtitleStyle.preserveLineBreaks',
        (src.subtitleStyle as { preserveLineBreaks?: unknown }).preserveLineBreaks,
        resolved.subtitleStyle.preserveLineBreaks,
        'Expected boolean.',
      );
    }

    const autoPauseVideoOnHover = asBoolean(
      (src.subtitleStyle as { autoPauseVideoOnHover?: unknown }).autoPauseVideoOnHover,
    );
    if (autoPauseVideoOnHover !== undefined) {
      resolved.subtitleStyle.autoPauseVideoOnHover = autoPauseVideoOnHover;
    } else if (
      (src.subtitleStyle as { autoPauseVideoOnHover?: unknown }).autoPauseVideoOnHover !== undefined
    ) {
      resolved.subtitleStyle.autoPauseVideoOnHover = fallbackSubtitleStyleAutoPauseVideoOnHover;
      warn(
        'subtitleStyle.autoPauseVideoOnHover',
        (src.subtitleStyle as { autoPauseVideoOnHover?: unknown }).autoPauseVideoOnHover,
        resolved.subtitleStyle.autoPauseVideoOnHover,
        'Expected boolean.',
      );
    }

    const autoPauseVideoOnYomitanPopup = asBoolean(
      (src.subtitleStyle as { autoPauseVideoOnYomitanPopup?: unknown })
        .autoPauseVideoOnYomitanPopup,
    );
    if (autoPauseVideoOnYomitanPopup !== undefined) {
      resolved.subtitleStyle.autoPauseVideoOnYomitanPopup = autoPauseVideoOnYomitanPopup;
    } else if (
      (src.subtitleStyle as { autoPauseVideoOnYomitanPopup?: unknown })
        .autoPauseVideoOnYomitanPopup !== undefined
    ) {
      resolved.subtitleStyle.autoPauseVideoOnYomitanPopup =
        fallbackSubtitleStyleAutoPauseVideoOnYomitanPopup;
      warn(
        'subtitleStyle.autoPauseVideoOnYomitanPopup',
        (src.subtitleStyle as { autoPauseVideoOnYomitanPopup?: unknown })
          .autoPauseVideoOnYomitanPopup,
        resolved.subtitleStyle.autoPauseVideoOnYomitanPopup,
        'Expected boolean.',
      );
    }

    const hoverTokenColor = asColor(
      (src.subtitleStyle as { hoverTokenColor?: unknown }).hoverTokenColor,
    );
    if (hoverTokenColor !== undefined) {
      resolved.subtitleStyle.hoverTokenColor = hoverTokenColor;
    } else if ((src.subtitleStyle as { hoverTokenColor?: unknown }).hoverTokenColor !== undefined) {
      resolved.subtitleStyle.hoverTokenColor = fallbackSubtitleStyleHoverTokenColor;
      warn(
        'subtitleStyle.hoverTokenColor',
        (src.subtitleStyle as { hoverTokenColor?: unknown }).hoverTokenColor,
        resolved.subtitleStyle.hoverTokenColor,
        'Expected hex color.',
      );
    }

    const subtitleStyleSource = src.subtitleStyle as {
      hoverBackground?: unknown;
      hoverTokenBackgroundColor?: unknown;
    };
    const rawHoverTokenBackgroundColor =
      subtitleStyleSource.hoverTokenBackgroundColor !== undefined
        ? subtitleStyleSource.hoverTokenBackgroundColor
        : subtitleStyleSource.hoverBackground;
    const hoverTokenBackgroundColor = asString(rawHoverTokenBackgroundColor);
    if (hoverTokenBackgroundColor !== undefined) {
      resolved.subtitleStyle.hoverTokenBackgroundColor = hoverTokenBackgroundColor;
    } else if (rawHoverTokenBackgroundColor !== undefined) {
      resolved.subtitleStyle.hoverTokenBackgroundColor =
        fallbackSubtitleStyleHoverTokenBackgroundColor;
      warn(
        'subtitleStyle.hoverTokenBackgroundColor',
        rawHoverTokenBackgroundColor,
        resolved.subtitleStyle.hoverTokenBackgroundColor,
        'Expected a CSS color value (hex, rgba/hsl/hsla, named color, or var()).',
      );
    }

    const nameMatchColor = asColor(
      (src.subtitleStyle as { nameMatchColor?: unknown }).nameMatchColor,
    );
    const nameMatchEnabled = asBoolean(
      (src.subtitleStyle as { nameMatchEnabled?: unknown }).nameMatchEnabled,
    );
    if (nameMatchEnabled !== undefined) {
      resolved.subtitleStyle.nameMatchEnabled = nameMatchEnabled;
    } else if (
      (src.subtitleStyle as { nameMatchEnabled?: unknown }).nameMatchEnabled !== undefined
    ) {
      resolved.subtitleStyle.nameMatchEnabled = fallbackSubtitleStyleNameMatchEnabled;
      warn(
        'subtitleStyle.nameMatchEnabled',
        (src.subtitleStyle as { nameMatchEnabled?: unknown }).nameMatchEnabled,
        resolved.subtitleStyle.nameMatchEnabled,
        'Expected boolean.',
      );
    }

    if (nameMatchColor !== undefined) {
      resolved.subtitleStyle.nameMatchColor = nameMatchColor;
    } else if ((src.subtitleStyle as { nameMatchColor?: unknown }).nameMatchColor !== undefined) {
      resolved.subtitleStyle.nameMatchColor = fallbackSubtitleStyleNameMatchColor;
      warn(
        'subtitleStyle.nameMatchColor',
        (src.subtitleStyle as { nameMatchColor?: unknown }).nameMatchColor,
        resolved.subtitleStyle.nameMatchColor,
        'Expected hex color.',
      );
    }

    const knownWordColor = asColor(
      (src.subtitleStyle as { knownWordColor?: unknown }).knownWordColor,
    );
    if (knownWordColor !== undefined) {
      resolved.subtitleStyle.knownWordColor = knownWordColor;
    } else if ((src.subtitleStyle as { knownWordColor?: unknown }).knownWordColor !== undefined) {
      resolved.subtitleStyle.knownWordColor = fallbackSubtitleStyleKnownWordColor;
      warn(
        'subtitleStyle.knownWordColor',
        (src.subtitleStyle as { knownWordColor?: unknown }).knownWordColor,
        resolved.subtitleStyle.knownWordColor,
        'Expected hex color.',
      );
    }

    const nPlusOneColor = asColor((src.subtitleStyle as { nPlusOneColor?: unknown }).nPlusOneColor);
    if (nPlusOneColor !== undefined) {
      resolved.subtitleStyle.nPlusOneColor = nPlusOneColor;
    } else if ((src.subtitleStyle as { nPlusOneColor?: unknown }).nPlusOneColor !== undefined) {
      resolved.subtitleStyle.nPlusOneColor = fallbackSubtitleStyleNPlusOneColor;
      warn(
        'subtitleStyle.nPlusOneColor',
        (src.subtitleStyle as { nPlusOneColor?: unknown }).nPlusOneColor,
        resolved.subtitleStyle.nPlusOneColor,
        'Expected hex color.',
      );
    }

    const frequencyDictionary = isObject(
      (src.subtitleStyle as { frequencyDictionary?: unknown }).frequencyDictionary,
    )
      ? ((src.subtitleStyle as { frequencyDictionary?: unknown }).frequencyDictionary as Record<
          string,
          unknown
        >)
      : {};
    const frequencyEnabled = asBoolean((frequencyDictionary as { enabled?: unknown }).enabled);
    if (frequencyEnabled !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.enabled = frequencyEnabled;
    } else if ((frequencyDictionary as { enabled?: unknown }).enabled !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.enabled = fallbackFrequencyDictionary.enabled;
      warn(
        'subtitleStyle.frequencyDictionary.enabled',
        (frequencyDictionary as { enabled?: unknown }).enabled,
        resolved.subtitleStyle.frequencyDictionary.enabled,
        'Expected boolean.',
      );
    }

    const sourcePath = asString((frequencyDictionary as { sourcePath?: unknown }).sourcePath);
    if (sourcePath !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.sourcePath = sourcePath;
    } else if ((frequencyDictionary as { sourcePath?: unknown }).sourcePath !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.sourcePath =
        fallbackFrequencyDictionary.sourcePath;
      warn(
        'subtitleStyle.frequencyDictionary.sourcePath',
        (frequencyDictionary as { sourcePath?: unknown }).sourcePath,
        resolved.subtitleStyle.frequencyDictionary.sourcePath,
        'Expected string.',
      );
    }

    const topX = asNumber((frequencyDictionary as { topX?: unknown }).topX);
    if (topX !== undefined && Number.isInteger(topX) && topX > 0) {
      resolved.subtitleStyle.frequencyDictionary.topX = Math.floor(topX);
    } else if ((frequencyDictionary as { topX?: unknown }).topX !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.topX = fallbackFrequencyDictionary.topX;
      warn(
        'subtitleStyle.frequencyDictionary.topX',
        (frequencyDictionary as { topX?: unknown }).topX,
        resolved.subtitleStyle.frequencyDictionary.topX,
        'Expected a positive integer.',
      );
    }

    const frequencyMode = frequencyDictionary.mode;
    if (frequencyMode === 'single' || frequencyMode === 'banded') {
      resolved.subtitleStyle.frequencyDictionary.mode = frequencyMode;
    } else if (frequencyMode !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.mode = fallbackFrequencyDictionary.mode;
      warn(
        'subtitleStyle.frequencyDictionary.mode',
        frequencyDictionary.mode,
        resolved.subtitleStyle.frequencyDictionary.mode,
        "Expected 'single' or 'banded'.",
      );
    }

    const frequencyMatchMode = (frequencyDictionary as { matchMode?: unknown }).matchMode;
    if (frequencyMatchMode === 'headword' || frequencyMatchMode === 'surface') {
      resolved.subtitleStyle.frequencyDictionary.matchMode = frequencyMatchMode;
    } else if (frequencyMatchMode !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.matchMode = fallbackFrequencyDictionary.matchMode;
      warn(
        'subtitleStyle.frequencyDictionary.matchMode',
        frequencyMatchMode,
        resolved.subtitleStyle.frequencyDictionary.matchMode,
        "Expected 'headword' or 'surface'.",
      );
    }

    const singleColor = asColor((frequencyDictionary as { singleColor?: unknown }).singleColor);
    if (singleColor !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.singleColor = singleColor;
    } else if ((frequencyDictionary as { singleColor?: unknown }).singleColor !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.singleColor =
        fallbackFrequencyDictionary.singleColor;
      warn(
        'subtitleStyle.frequencyDictionary.singleColor',
        (frequencyDictionary as { singleColor?: unknown }).singleColor,
        resolved.subtitleStyle.frequencyDictionary.singleColor,
        'Expected hex color.',
      );
    }

    const bandedColors = asFrequencyBandedColors(
      (frequencyDictionary as { bandedColors?: unknown }).bandedColors,
    );
    if (bandedColors !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.bandedColors = bandedColors;
    } else if ((frequencyDictionary as { bandedColors?: unknown }).bandedColors !== undefined) {
      resolved.subtitleStyle.frequencyDictionary.bandedColors =
        fallbackFrequencyDictionary.bandedColors;
      warn(
        'subtitleStyle.frequencyDictionary.bandedColors',
        (frequencyDictionary as { bandedColors?: unknown }).bandedColors,
        resolved.subtitleStyle.frequencyDictionary.bandedColors,
        'Expected an array of five hex colors.',
      );
    }
  }

  if (isObject(src.subtitleSidebar)) {
    const fallback = { ...resolved.subtitleSidebar };
    resolved.subtitleSidebar = {
      ...resolved.subtitleSidebar,
      ...(src.subtitleSidebar as ResolvedConfig['subtitleSidebar']),
    };

    const enabled = asBoolean((src.subtitleSidebar as { enabled?: unknown }).enabled);
    if (enabled !== undefined) {
      resolved.subtitleSidebar.enabled = enabled;
    } else if ((src.subtitleSidebar as { enabled?: unknown }).enabled !== undefined) {
      resolved.subtitleSidebar.enabled = fallback.enabled;
      warn(
        'subtitleSidebar.enabled',
        (src.subtitleSidebar as { enabled?: unknown }).enabled,
        resolved.subtitleSidebar.enabled,
        'Expected boolean.',
      );
    }

    const autoOpen = asBoolean((src.subtitleSidebar as { autoOpen?: unknown }).autoOpen);
    if (autoOpen !== undefined) {
      resolved.subtitleSidebar.autoOpen = autoOpen;
    } else if ((src.subtitleSidebar as { autoOpen?: unknown }).autoOpen !== undefined) {
      resolved.subtitleSidebar.autoOpen = fallback.autoOpen;
      warn(
        'subtitleSidebar.autoOpen',
        (src.subtitleSidebar as { autoOpen?: unknown }).autoOpen,
        resolved.subtitleSidebar.autoOpen,
        'Expected boolean.',
      );
    }

    const layout = asString((src.subtitleSidebar as { layout?: unknown }).layout);
    if (layout === 'overlay' || layout === 'embedded') {
      resolved.subtitleSidebar.layout = layout;
    } else if ((src.subtitleSidebar as { layout?: unknown }).layout !== undefined) {
      resolved.subtitleSidebar.layout = fallback.layout;
      warn(
        'subtitleSidebar.layout',
        (src.subtitleSidebar as { layout?: unknown }).layout,
        resolved.subtitleSidebar.layout,
        'Expected "overlay" or "embedded".',
      );
    }

    const pauseVideoOnHover = asBoolean(
      (src.subtitleSidebar as { pauseVideoOnHover?: unknown }).pauseVideoOnHover,
    );
    if (pauseVideoOnHover !== undefined) {
      resolved.subtitleSidebar.pauseVideoOnHover = pauseVideoOnHover;
    } else if (
      (src.subtitleSidebar as { pauseVideoOnHover?: unknown }).pauseVideoOnHover !== undefined
    ) {
      resolved.subtitleSidebar.pauseVideoOnHover = fallback.pauseVideoOnHover;
      warn(
        'subtitleSidebar.pauseVideoOnHover',
        (src.subtitleSidebar as { pauseVideoOnHover?: unknown }).pauseVideoOnHover,
        resolved.subtitleSidebar.pauseVideoOnHover,
        'Expected boolean.',
      );
    }

    const autoScroll = asBoolean((src.subtitleSidebar as { autoScroll?: unknown }).autoScroll);
    if (autoScroll !== undefined) {
      resolved.subtitleSidebar.autoScroll = autoScroll;
    } else if ((src.subtitleSidebar as { autoScroll?: unknown }).autoScroll !== undefined) {
      resolved.subtitleSidebar.autoScroll = fallback.autoScroll;
      warn(
        'subtitleSidebar.autoScroll',
        (src.subtitleSidebar as { autoScroll?: unknown }).autoScroll,
        resolved.subtitleSidebar.autoScroll,
        'Expected boolean.',
      );
    }

    const toggleKey = asString((src.subtitleSidebar as { toggleKey?: unknown }).toggleKey);
    if (toggleKey !== undefined) {
      resolved.subtitleSidebar.toggleKey = toggleKey;
    } else if ((src.subtitleSidebar as { toggleKey?: unknown }).toggleKey !== undefined) {
      resolved.subtitleSidebar.toggleKey = fallback.toggleKey;
      warn(
        'subtitleSidebar.toggleKey',
        (src.subtitleSidebar as { toggleKey?: unknown }).toggleKey,
        resolved.subtitleSidebar.toggleKey,
        'Expected string.',
      );
    }

    const maxWidth = asNumber((src.subtitleSidebar as { maxWidth?: unknown }).maxWidth);
    if (maxWidth !== undefined && maxWidth > 0) {
      resolved.subtitleSidebar.maxWidth = Math.floor(maxWidth);
    } else if ((src.subtitleSidebar as { maxWidth?: unknown }).maxWidth !== undefined) {
      resolved.subtitleSidebar.maxWidth = fallback.maxWidth;
      warn(
        'subtitleSidebar.maxWidth',
        (src.subtitleSidebar as { maxWidth?: unknown }).maxWidth,
        resolved.subtitleSidebar.maxWidth,
        'Expected positive number.',
      );
    }

    const opacity = asNumber((src.subtitleSidebar as { opacity?: unknown }).opacity);
    if (opacity !== undefined && opacity >= 0 && opacity <= 1) {
      resolved.subtitleSidebar.opacity = opacity;
    } else if ((src.subtitleSidebar as { opacity?: unknown }).opacity !== undefined) {
      resolved.subtitleSidebar.opacity = fallback.opacity;
      warn(
        'subtitleSidebar.opacity',
        (src.subtitleSidebar as { opacity?: unknown }).opacity,
        resolved.subtitleSidebar.opacity,
        'Expected number between 0 and 1.',
      );
    }

    const hexColorFields = ['textColor', 'timestampColor', 'activeLineColor'] as const;
    for (const field of hexColorFields) {
      const value = asColor((src.subtitleSidebar as Record<string, unknown>)[field]);
      if (value !== undefined) {
        resolved.subtitleSidebar[field] = value;
      } else if ((src.subtitleSidebar as Record<string, unknown>)[field] !== undefined) {
        resolved.subtitleSidebar[field] = fallback[field];
        warn(
          `subtitleSidebar.${field}`,
          (src.subtitleSidebar as Record<string, unknown>)[field],
          resolved.subtitleSidebar[field],
          'Expected hex color.',
        );
      }
    }

    const cssColorFields = [
      'backgroundColor',
      'activeLineBackgroundColor',
      'hoverLineBackgroundColor',
    ] as const;
    for (const field of cssColorFields) {
      const value = asCssColor((src.subtitleSidebar as Record<string, unknown>)[field]);
      if (value !== undefined) {
        resolved.subtitleSidebar[field] = value;
      } else if ((src.subtitleSidebar as Record<string, unknown>)[field] !== undefined) {
        resolved.subtitleSidebar[field] = fallback[field];
        warn(
          `subtitleSidebar.${field}`,
          (src.subtitleSidebar as Record<string, unknown>)[field],
          resolved.subtitleSidebar[field],
          'Expected valid CSS color.',
        );
      }
    }

    const fontFamily = asString((src.subtitleSidebar as { fontFamily?: unknown }).fontFamily);
    if (fontFamily !== undefined && fontFamily.trim().length > 0) {
      resolved.subtitleSidebar.fontFamily = fontFamily.trim();
    } else if ((src.subtitleSidebar as { fontFamily?: unknown }).fontFamily !== undefined) {
      resolved.subtitleSidebar.fontFamily = fallback.fontFamily;
      warn(
        'subtitleSidebar.fontFamily',
        (src.subtitleSidebar as { fontFamily?: unknown }).fontFamily,
        resolved.subtitleSidebar.fontFamily,
        'Expected non-empty string.',
      );
    }

    const fontSize = asNumber((src.subtitleSidebar as { fontSize?: unknown }).fontSize);
    if (fontSize !== undefined && fontSize > 0) {
      resolved.subtitleSidebar.fontSize = fontSize;
    } else if ((src.subtitleSidebar as { fontSize?: unknown }).fontSize !== undefined) {
      resolved.subtitleSidebar.fontSize = fallback.fontSize;
      warn(
        'subtitleSidebar.fontSize',
        (src.subtitleSidebar as { fontSize?: unknown }).fontSize,
        resolved.subtitleSidebar.fontSize,
        'Expected positive number.',
      );
    }
  }
}
