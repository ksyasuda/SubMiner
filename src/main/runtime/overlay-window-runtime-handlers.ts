import {
  createCreateMainWindowHandler,
  createCreateModalWindowHandler,
  createCreateOverlayWindowHandler,
} from './overlay-window-factory';
import {
  createBuildCreateMainWindowMainDepsHandler,
  createBuildCreateModalWindowMainDepsHandler,
  createBuildCreateOverlayWindowMainDepsHandler,
} from './overlay-window-factory-main-deps';

type CreateOverlayWindowMainDeps<TWindow> = Parameters<
  typeof createBuildCreateOverlayWindowMainDepsHandler<TWindow>
>[0];

export function createOverlayWindowRuntimeHandlers<TWindow>(deps: {
  createOverlayWindowDeps: CreateOverlayWindowMainDeps<TWindow>;
  getMainWindow: () => TWindow | null;
  isWindowDestroyed: (window: TWindow) => boolean;
  setMainWindow: (window: TWindow | null) => void;
  setModalWindow: (window: TWindow | null) => void;
}) {
  const createOverlayWindow = createCreateOverlayWindowHandler<TWindow>(
    createBuildCreateOverlayWindowMainDepsHandler<TWindow>(deps.createOverlayWindowDeps)(),
  );
  const createMainWindow = createCreateMainWindowHandler<TWindow>(
    createBuildCreateMainWindowMainDepsHandler<TWindow>({
      getMainWindow: () => deps.getMainWindow(),
      isWindowDestroyed: (window) => deps.isWindowDestroyed(window),
      createOverlayWindow: (kind) => createOverlayWindow(kind),
      setMainWindow: (window) => deps.setMainWindow(window),
    })(),
  );
  const createModalWindow = createCreateModalWindowHandler<TWindow>(
    createBuildCreateModalWindowMainDepsHandler<TWindow>({
      createOverlayWindow: (kind) => createOverlayWindow(kind),
      setModalWindow: (window) => deps.setModalWindow(window),
    })(),
  );

  return {
    createOverlayWindow,
    createMainWindow,
    createModalWindow,
  };
}
