import { PartOfSpeech as LegacyPartOfSpeech } from './types';
import type { AnkiConnectConfig } from './types/anki';
import type { ConfigValidationWarning, RawConfig, ResolvedConfig } from './types/config';
import type { JimakuConfig, YoutubePickerOpenPayload } from './types/integrations';
import type { ElectronAPI, MpvClient, OverlayContentMeasurement } from './types/runtime';
import type { RuntimeOptionId, RuntimeOptionValue } from './types/runtime-options';
import { PartOfSpeech, type SubtitleSidebarSnapshot } from './types/subtitle';

type Assert<T extends true> = T;
type IsAssignable<From, To> = [From] extends [To] ? true : false;

const runtimeEntryPointMatchesLegacyBarrel = LegacyPartOfSpeech === PartOfSpeech;
void runtimeEntryPointMatchesLegacyBarrel;

type SubtitleEnumStillCompatible = Assert<
  IsAssignable<typeof PartOfSpeech, typeof LegacyPartOfSpeech>
>;

type ConfigEntryPointContracts = [
  RawConfig,
  ResolvedConfig,
  ConfigValidationWarning,
  RuntimeOptionId,
  RuntimeOptionValue,
];

type IntegrationEntryPointContracts = [AnkiConnectConfig, JimakuConfig, YoutubePickerOpenPayload];

type RuntimeEntryPointContracts = [
  OverlayContentMeasurement,
  ElectronAPI,
  MpvClient,
  SubtitleSidebarSnapshot,
];

void (null as unknown as SubtitleEnumStillCompatible);
void (null as unknown as ConfigEntryPointContracts);
void (null as unknown as IntegrationEntryPointContracts);
void (null as unknown as RuntimeEntryPointContracts);
