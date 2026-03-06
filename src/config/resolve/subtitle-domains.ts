import { ResolvedConfig } from '../../types';
import { ResolveContext } from './context';
import {
  asBoolean,
  asColor,
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
    const mode = src.youtubeSubgen.mode;
    if (mode === 'automatic' || mode === 'preprocess' || mode === 'off') {
      resolved.youtubeSubgen.mode = mode;
    } else if (mode !== undefined) {
      warn(
        'youtubeSubgen.mode',
        mode,
        resolved.youtubeSubgen.mode,
        'Expected automatic, preprocess, or off.',
      );
    }

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

    if (Array.isArray(src.youtubeSubgen.primarySubLanguages)) {
      resolved.youtubeSubgen.primarySubLanguages = src.youtubeSubgen.primarySubLanguages.filter(
        (item): item is string => typeof item === 'string',
      );
    } else if (src.youtubeSubgen.primarySubLanguages !== undefined) {
      warn(
        'youtubeSubgen.primarySubLanguages',
        src.youtubeSubgen.primarySubLanguages,
        resolved.youtubeSubgen.primarySubLanguages,
        'Expected string array.',
      );
    }
  }

  if (isObject(src.subtitleStyle)) {
    const fallbackSubtitleStyleEnableJlpt = resolved.subtitleStyle.enableJlpt;
    const fallbackSubtitleStylePreserveLineBreaks = resolved.subtitleStyle.preserveLineBreaks;
    const fallbackSubtitleStyleAutoPauseVideoOnHover = resolved.subtitleStyle.autoPauseVideoOnHover;
    const fallbackSubtitleStyleAutoPauseVideoOnYomitanPopup =
      resolved.subtitleStyle.autoPauseVideoOnYomitanPopup;
    const fallbackSubtitleStyleHoverTokenColor = resolved.subtitleStyle.hoverTokenColor;
    const fallbackSubtitleStyleHoverTokenBackgroundColor =
      resolved.subtitleStyle.hoverTokenBackgroundColor;
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

    const hoverTokenBackgroundColor = asString(
      (src.subtitleStyle as { hoverTokenBackgroundColor?: unknown }).hoverTokenBackgroundColor,
    );
    if (hoverTokenBackgroundColor !== undefined) {
      resolved.subtitleStyle.hoverTokenBackgroundColor = hoverTokenBackgroundColor;
    } else if (
      (src.subtitleStyle as { hoverTokenBackgroundColor?: unknown }).hoverTokenBackgroundColor !==
      undefined
    ) {
      resolved.subtitleStyle.hoverTokenBackgroundColor =
        fallbackSubtitleStyleHoverTokenBackgroundColor;
      warn(
        'subtitleStyle.hoverTokenBackgroundColor',
        (src.subtitleStyle as { hoverTokenBackgroundColor?: unknown }).hoverTokenBackgroundColor,
        resolved.subtitleStyle.hoverTokenBackgroundColor,
        'Expected a CSS color value (hex, rgba/hsl/hsla, named color, or var()).',
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
}
