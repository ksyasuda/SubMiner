import type { ComposerInputs } from './contracts';
import type { IpcRuntimeComposerOptions } from './ipc-runtime-composer';
import type { JellyfinRemoteComposerOptions } from './jellyfin-remote-composer';
import type { MpvRuntimeComposerOptions } from './mpv-runtime-composer';
import type { AnilistSetupComposerOptions } from './anilist-setup-composer';

type Assert<T extends true> = T;
type IsAssignable<From, To> = [From] extends [To] ? true : false;

type FakeMpvClient = {
  on: (...args: unknown[]) => unknown;
  connect: () => void;
};

type FakeTokenizerDeps = { isKnownWord: (text: string) => boolean };
type FakeTokenizedSubtitle = { text: string };

type RequiredAnilistSetupInputKeys = keyof ComposerInputs<AnilistSetupComposerOptions>;
type RequiredJellyfinInputKeys = keyof ComposerInputs<JellyfinRemoteComposerOptions>;
type RequiredIpcInputKeys = keyof ComposerInputs<IpcRuntimeComposerOptions>;
type RequiredMpvInputKeys = keyof ComposerInputs<
  MpvRuntimeComposerOptions<FakeMpvClient, FakeTokenizerDeps, FakeTokenizedSubtitle>
>;

const contractAssertions = [
  true as Assert<IsAssignable<'notifyDeps', RequiredAnilistSetupInputKeys>>,
  true as Assert<IsAssignable<'getMpvClient', RequiredJellyfinInputKeys>>,
  true as Assert<IsAssignable<'registration', RequiredIpcInputKeys>>,
  true as Assert<IsAssignable<'tokenizer', RequiredMpvInputKeys>>,
];
void contractAssertions;

// @ts-expect-error missing required notifyDeps should fail compile-time contract
const anilistMissingRequired: AnilistSetupComposerOptions = {
  consumeTokenDeps: {} as AnilistSetupComposerOptions['consumeTokenDeps'],
  handleProtocolDeps: {} as AnilistSetupComposerOptions['handleProtocolDeps'],
  registerProtocolClientDeps: {} as AnilistSetupComposerOptions['registerProtocolClientDeps'],
};

// @ts-expect-error missing required getMpvClient should fail compile-time contract
const jellyfinMissingRequired: JellyfinRemoteComposerOptions = {
  getConfiguredSession: {} as JellyfinRemoteComposerOptions['getConfiguredSession'],
  getClientInfo: {} as JellyfinRemoteComposerOptions['getClientInfo'],
  getJellyfinConfig: {} as JellyfinRemoteComposerOptions['getJellyfinConfig'],
  playJellyfinItem: {} as JellyfinRemoteComposerOptions['playJellyfinItem'],
  logWarn: {} as JellyfinRemoteComposerOptions['logWarn'],
  sendMpvCommand: {} as JellyfinRemoteComposerOptions['sendMpvCommand'],
  jellyfinTicksToSeconds: {} as JellyfinRemoteComposerOptions['jellyfinTicksToSeconds'],
  getActivePlayback: {} as JellyfinRemoteComposerOptions['getActivePlayback'],
  clearActivePlayback: {} as JellyfinRemoteComposerOptions['clearActivePlayback'],
  getSession: {} as JellyfinRemoteComposerOptions['getSession'],
  getNow: {} as JellyfinRemoteComposerOptions['getNow'],
  getLastProgressAtMs: {} as JellyfinRemoteComposerOptions['getLastProgressAtMs'],
  setLastProgressAtMs: {} as JellyfinRemoteComposerOptions['setLastProgressAtMs'],
  progressIntervalMs: 3000,
  ticksPerSecond: 10_000_000,
  logDebug: {} as JellyfinRemoteComposerOptions['logDebug'],
};

// @ts-expect-error missing required registration should fail compile-time contract
const ipcMissingRequired: IpcRuntimeComposerOptions = {
  mpvCommandMainDeps: {} as IpcRuntimeComposerOptions['mpvCommandMainDeps'],
  handleMpvCommandFromIpcRuntime: {} as IpcRuntimeComposerOptions['handleMpvCommandFromIpcRuntime'],
  runSubsyncManualFromIpc: {} as IpcRuntimeComposerOptions['runSubsyncManualFromIpc'],
};

// @ts-expect-error missing required tokenizer should fail compile-time contract
const mpvMissingRequired: MpvRuntimeComposerOptions<
  FakeMpvClient,
  FakeTokenizerDeps,
  FakeTokenizedSubtitle
> = {
  bindMpvMainEventHandlersMainDeps: {} as MpvRuntimeComposerOptions<
    FakeMpvClient,
    FakeTokenizerDeps,
    FakeTokenizedSubtitle
  >['bindMpvMainEventHandlersMainDeps'],
  mpvClientRuntimeServiceFactoryMainDeps: {} as MpvRuntimeComposerOptions<
    FakeMpvClient,
    FakeTokenizerDeps,
    FakeTokenizedSubtitle
  >['mpvClientRuntimeServiceFactoryMainDeps'],
  updateMpvSubtitleRenderMetricsMainDeps: {} as MpvRuntimeComposerOptions<
    FakeMpvClient,
    FakeTokenizerDeps,
    FakeTokenizedSubtitle
  >['updateMpvSubtitleRenderMetricsMainDeps'],
  warmups: {} as MpvRuntimeComposerOptions<
    FakeMpvClient,
    FakeTokenizerDeps,
    FakeTokenizedSubtitle
  >['warmups'],
};

void anilistMissingRequired;
void jellyfinMissingRequired;
void ipcMissingRequired;
void mpvMissingRequired;
