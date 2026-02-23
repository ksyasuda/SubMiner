import type { OverlayShortcutRuntimeServiceInput } from '../overlay-shortcuts-runtime';

export function createBuildOverlayShortcutsRuntimeMainDepsHandler(
  deps: OverlayShortcutRuntimeServiceInput,
) {
  return (): OverlayShortcutRuntimeServiceInput => ({
    getConfiguredShortcuts: () => deps.getConfiguredShortcuts(),
    getShortcutsRegistered: () => deps.getShortcutsRegistered(),
    setShortcutsRegistered: (registered: boolean) => deps.setShortcutsRegistered(registered),
    isOverlayRuntimeInitialized: () => deps.isOverlayRuntimeInitialized(),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    openRuntimeOptionsPalette: () => deps.openRuntimeOptionsPalette(),
    openJimaku: () => deps.openJimaku(),
    markAudioCard: () => deps.markAudioCard(),
    copySubtitleMultiple: (timeoutMs: number) => deps.copySubtitleMultiple(timeoutMs),
    copySubtitle: () => deps.copySubtitle(),
    toggleSecondarySubMode: () => deps.toggleSecondarySubMode(),
    updateLastCardFromClipboard: () => deps.updateLastCardFromClipboard(),
    triggerFieldGrouping: () => deps.triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
    mineSentenceCard: () => deps.mineSentenceCard(),
    mineSentenceMultiple: (timeoutMs: number) => deps.mineSentenceMultiple(timeoutMs),
    cancelPendingMultiCopy: () => deps.cancelPendingMultiCopy(),
    cancelPendingMineSentenceMultiple: () => deps.cancelPendingMineSentenceMultiple(),
  });
}
