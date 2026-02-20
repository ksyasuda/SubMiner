import {
  createCreateInvisibleWindowHandler,
  createCreateMainWindowHandler,
  createCreateOverlayWindowHandler,
} from './overlay-window-factory';
import {
  createBuildCreateInvisibleWindowMainDepsHandler,
  createBuildCreateMainWindowMainDepsHandler,
  createBuildCreateOverlayWindowMainDepsHandler,
} from './overlay-window-factory-main-deps';

type CreateOverlayWindowMainDeps<TWindow> = Parameters<
  typeof createBuildCreateOverlayWindowMainDepsHandler<TWindow>
>[0];

export function createOverlayWindowRuntimeHandlers<TWindow>(deps: {
  createOverlayWindowDeps: CreateOverlayWindowMainDeps<TWindow>;
  setMainWindow: (window: TWindow | null) => void;
  setInvisibleWindow: (window: TWindow | null) => void;
}) {
  const createOverlayWindow = createCreateOverlayWindowHandler<TWindow>(
    createBuildCreateOverlayWindowMainDepsHandler<TWindow>(deps.createOverlayWindowDeps)(),
  );
  const createMainWindow = createCreateMainWindowHandler<TWindow>(
    createBuildCreateMainWindowMainDepsHandler<TWindow>({
      createOverlayWindow: (kind) => createOverlayWindow(kind),
      setMainWindow: (window) => deps.setMainWindow(window),
    })(),
  );
  const createInvisibleWindow = createCreateInvisibleWindowHandler<TWindow>(
    createBuildCreateInvisibleWindowMainDepsHandler<TWindow>({
      createOverlayWindow: (kind) => createOverlayWindow(kind),
      setInvisibleWindow: (window) => deps.setInvisibleWindow(window),
    })(),
  );

  return {
    createOverlayWindow,
    createMainWindow,
    createInvisibleWindow,
  };
}
