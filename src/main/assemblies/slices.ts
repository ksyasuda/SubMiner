import type { RegisterIpcRuntimeServicesParams } from '../ipc-runtime';
import type { CliCommandContextMainDeps } from '../runtime/cli-command-context-factory';

/** Renderer IPC handlers, minus the two the IPC composer builds itself. */
export type IpcHandlerDeps = Omit<
  RegisterIpcRuntimeServicesParams['mainDeps'],
  'handleMpvCommand' | 'runSubsyncManual'
>;

/**
 * The IPC handlers one assembly provides. `Required` because most handlers are optional on
 * the registration params, and a slice must not drop one silently.
 */
export type IpcSlice<K extends keyof IpcHandlerDeps> = Required<Pick<IpcHandlerDeps, K>>;

/** The CLI command handlers one assembly provides. */
export type CliSlice<K extends keyof CliCommandContextMainDeps> = Required<
  Pick<CliCommandContextMainDeps, K>
>;
