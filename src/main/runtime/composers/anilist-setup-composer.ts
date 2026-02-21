import {
  createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler,
  createBuildHandleAnilistSetupProtocolUrlMainDepsHandler,
  createBuildNotifyAnilistSetupMainDepsHandler,
  createBuildRegisterSubminerProtocolClientMainDepsHandler,
  createConsumeAnilistSetupTokenFromUrlHandler,
  createHandleAnilistSetupProtocolUrlHandler,
  createNotifyAnilistSetupHandler,
  createRegisterSubminerProtocolClientHandler,
} from '../domains/anilist';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type NotifyHandler = ReturnType<typeof createNotifyAnilistSetupHandler>;
type ConsumeHandler = ReturnType<typeof createConsumeAnilistSetupTokenFromUrlHandler>;
type HandleProtocolHandler = ReturnType<typeof createHandleAnilistSetupProtocolUrlHandler>;
type RegisterClientHandler = ReturnType<typeof createRegisterSubminerProtocolClientHandler>;

export type AnilistSetupComposerOptions = ComposerInputs<{
  notifyDeps: Parameters<typeof createBuildNotifyAnilistSetupMainDepsHandler>[0];
  consumeTokenDeps: Parameters<typeof createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler>[0];
  handleProtocolDeps: Parameters<typeof createBuildHandleAnilistSetupProtocolUrlMainDepsHandler>[0];
  registerProtocolClientDeps: Parameters<
    typeof createBuildRegisterSubminerProtocolClientMainDepsHandler
  >[0];
}>;

export type AnilistSetupComposerResult = ComposerOutputs<{
  notifyAnilistSetup: NotifyHandler;
  consumeAnilistSetupTokenFromUrl: ConsumeHandler;
  handleAnilistSetupProtocolUrl: HandleProtocolHandler;
  registerSubminerProtocolClient: RegisterClientHandler;
}>;

export function composeAnilistSetupHandlers(
  options: AnilistSetupComposerOptions,
): AnilistSetupComposerResult {
  const notifyAnilistSetup = createNotifyAnilistSetupHandler(
    createBuildNotifyAnilistSetupMainDepsHandler(options.notifyDeps)(),
  );
  const consumeAnilistSetupTokenFromUrl = createConsumeAnilistSetupTokenFromUrlHandler(
    createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler(options.consumeTokenDeps)(),
  );
  const handleAnilistSetupProtocolUrl = createHandleAnilistSetupProtocolUrlHandler(
    createBuildHandleAnilistSetupProtocolUrlMainDepsHandler(options.handleProtocolDeps)(),
  );
  const registerSubminerProtocolClient = createRegisterSubminerProtocolClientHandler(
    createBuildRegisterSubminerProtocolClientMainDepsHandler(options.registerProtocolClientDeps)(),
  );

  return {
    notifyAnilistSetup,
    consumeAnilistSetupTokenFromUrl,
    handleAnilistSetupProtocolUrl,
    registerSubminerProtocolClient,
  };
}
