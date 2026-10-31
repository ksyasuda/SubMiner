import {
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
  notifyDeps: Parameters<typeof createNotifyAnilistSetupHandler>[0];
  consumeTokenDeps: Parameters<typeof createConsumeAnilistSetupTokenFromUrlHandler>[0];
  handleProtocolDeps: Parameters<typeof createHandleAnilistSetupProtocolUrlHandler>[0];
  registerProtocolClientDeps: Parameters<typeof createRegisterSubminerProtocolClientHandler>[0];
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
  const notifyAnilistSetup = createNotifyAnilistSetupHandler(options.notifyDeps);
  const consumeAnilistSetupTokenFromUrl = createConsumeAnilistSetupTokenFromUrlHandler(
    options.consumeTokenDeps,
  );
  const handleAnilistSetupProtocolUrl = createHandleAnilistSetupProtocolUrlHandler(
    options.handleProtocolDeps,
  );
  const registerSubminerProtocolClient = createRegisterSubminerProtocolClientHandler(
    options.registerProtocolClientDeps,
  );

  return {
    notifyAnilistSetup,
    consumeAnilistSetupTokenFromUrl,
    handleAnilistSetupProtocolUrl,
    registerSubminerProtocolClient,
  };
}
