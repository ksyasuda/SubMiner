import type { RendererState } from './state';
import type { RendererDom } from './utils/dom';
import type { PlatformInfo } from './utils/platform';

export type RendererContext = {
  dom: RendererDom;
  platform: PlatformInfo;
  state: RendererState;
};

export type ModalStateReader = {
  isAnySettingsModalOpen: () => boolean;
  isAnyModalOpen: () => boolean;
};
