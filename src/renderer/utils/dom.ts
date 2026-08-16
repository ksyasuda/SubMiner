export type RendererDom = {
  subtitleRoot: HTMLElement;
  subtitleContainer: HTMLElement;
  overlay: HTMLElement;
  overlayNotificationStack: HTMLDivElement;
  overlayNotificationHistory: HTMLElement;
  controllerStatusToast: HTMLDivElement;
  overlayErrorToast: HTMLDivElement;
  secondarySubContainer: HTMLElement;
  secondarySubRoot: HTMLElement;

  jimakuModal: HTMLDivElement;
  jimakuTitleInput: HTMLInputElement;
  jimakuSeasonInput: HTMLInputElement;
  jimakuEpisodeInput: HTMLInputElement;
  jimakuSearchButton: HTMLButtonElement;
  jimakuCloseButton: HTMLButtonElement;
  jimakuStatus: HTMLDivElement;
  jimakuEntriesSection: HTMLDivElement;
  jimakuEntriesList: HTMLUListElement;
  jimakuFilesSection: HTMLDivElement;
  jimakuFilesList: HTMLUListElement;
  jimakuBroadenButton: HTMLButtonElement;

  tsukihimeModal: HTMLDivElement;
  tsukihimeTitleInput: HTMLInputElement;
  tsukihimeSeasonInput: HTMLInputElement;
  tsukihimeEpisodeInput: HTMLInputElement;
  tsukihimeSearchButton: HTMLButtonElement;
  tsukihimeCloseButton: HTMLButtonElement;
  tsukihimeTabSecondaryButton: HTMLButtonElement;
  tsukihimeTabPrimaryButton: HTMLButtonElement;
  tsukihimeStatus: HTMLDivElement;
  tsukihimeEntriesSection: HTMLDivElement;
  tsukihimeEntriesList: HTMLUListElement;
  tsukihimeFilesSection: HTMLDivElement;
  tsukihimeFilesList: HTMLUListElement;

  youtubePickerModal: HTMLDivElement;
  youtubePickerTitle: HTMLDivElement;
  youtubePickerPrimarySelect: HTMLSelectElement;
  youtubePickerSecondarySelect: HTMLSelectElement;
  youtubePickerContinueButton: HTMLButtonElement;
  youtubePickerCloseButton: HTMLButtonElement;
  youtubePickerStatus: HTMLDivElement;
  youtubePickerTracks: HTMLUListElement;

  kikuModal: HTMLDivElement;
  kikuCard1: HTMLDivElement;
  kikuCard2: HTMLDivElement;
  kikuCard1Expression: HTMLDivElement;
  kikuCard2Expression: HTMLDivElement;
  kikuCard1Sentence: HTMLDivElement;
  kikuCard2Sentence: HTMLDivElement;
  kikuCard1Meta: HTMLDivElement;
  kikuCard2Meta: HTMLDivElement;
  kikuConfirmButton: HTMLButtonElement;
  kikuCancelButton: HTMLButtonElement;
  kikuDeleteDuplicateCheckbox: HTMLInputElement;
  kikuSelectionStep: HTMLDivElement;
  kikuPreviewStep: HTMLDivElement;
  kikuPreviewJson: HTMLPreElement;
  kikuPreviewCompactButton: HTMLButtonElement;
  kikuPreviewFullButton: HTMLButtonElement;
  kikuPreviewError: HTMLDivElement;
  kikuBackButton: HTMLButtonElement;
  kikuFinalConfirmButton: HTMLButtonElement;
  kikuFinalCancelButton: HTMLButtonElement;
  kikuHint: HTMLDivElement;

  runtimeOptionsModal: HTMLDivElement;
  runtimeOptionsClose: HTMLButtonElement;
  runtimeOptionsList: HTMLUListElement;
  runtimeOptionsStatus: HTMLDivElement;

  characterDictionaryModal: HTMLDivElement;
  characterDictionaryClose: HTMLButtonElement;
  characterDictionaryOverrideTab: HTMLButtonElement;
  characterDictionaryManageTab: HTMLButtonElement;
  characterDictionarySummary: HTMLDivElement;
  characterDictionarySearchPanel: HTMLDivElement;
  characterDictionarySearchInput: HTMLInputElement;
  characterDictionarySearchButton: HTMLButtonElement;
  characterDictionaryCurrent: HTMLDivElement;
  characterDictionaryCandidates: HTMLUListElement;
  characterDictionaryManagerPanel: HTMLDivElement;
  characterDictionaryManagedEntries: HTMLUListElement;
  characterDictionaryStatus: HTMLDivElement;

  subsyncModal: HTMLDivElement;
  subsyncCloseButton: HTMLButtonElement;
  subsyncEngineAlass: HTMLInputElement;
  subsyncEngineFfsubsync: HTMLInputElement;
  subsyncReferenceLabel: HTMLLabelElement;
  subsyncReferenceSelect: HTMLSelectElement;
  subsyncTargetLabel: HTMLLabelElement;
  subsyncTargetSelect: HTMLSelectElement;
  subsyncRunButton: HTMLButtonElement;
  subsyncStatus: HTMLDivElement;

  controllerSelectModal: HTMLDivElement;
  controllerSelectClose: HTMLButtonElement;
  controllerSelectPicker: HTMLSelectElement;
  controllerSelectSummary: HTMLDivElement;
  controllerSelectStatus: HTMLDivElement;
  controllerConfigList: HTMLDivElement;
  controllerSelectSave: HTMLButtonElement;

  controllerDebugModal: HTMLDivElement;
  controllerDebugClose: HTMLButtonElement;
  controllerDebugCopy: HTMLButtonElement;
  controllerDebugToast: HTMLDivElement;
  controllerDebugStatus: HTMLDivElement;
  controllerDebugSummary: HTMLDivElement;
  controllerDebugAxes: HTMLPreElement;
  controllerDebugButtons: HTMLPreElement;
  controllerDebugButtonIndices: HTMLPreElement;
  subtitleSidebarModal: HTMLDivElement;
  subtitleSidebarContent: HTMLDivElement;
  subtitleSidebarClose: HTMLButtonElement;
  subtitleSidebarStatus: HTMLDivElement;
  subtitleSidebarList: HTMLUListElement;

  changelogModal: HTMLDivElement;
  changelogClose: HTMLButtonElement;
  changelogRefresh: HTMLButtonElement;
  changelogInstalled: HTMLSpanElement;
  changelogSource: HTMLSpanElement;
  changelogWarning: HTMLDivElement;
  changelogStatus: HTMLDivElement;
  changelogList: HTMLDivElement;

  sessionHelpModal: HTMLDivElement;
  sessionHelpClose: HTMLButtonElement;
  sessionHelpShortcut: HTMLDivElement;
  sessionHelpWarning: HTMLDivElement;
  sessionHelpStatus: HTMLDivElement;
  sessionHelpFilter: HTMLInputElement;
  sessionHelpContent: HTMLDivElement;

  playlistBrowserModal: HTMLDivElement;
  playlistBrowserTitle: HTMLDivElement;
  playlistBrowserStatus: HTMLDivElement;
  playlistBrowserDirectoryList: HTMLUListElement;
  playlistBrowserPlaylistList: HTMLUListElement;
  playlistBrowserClose: HTMLButtonElement;

  animeBrowserModal: HTMLDivElement;
  animeBrowserClose: HTMLButtonElement;
  animeBrowserFrame: HTMLIFrameElement;
};

function getRequiredElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing required DOM element #${id}`);
  }
  return element as T;
}

export function resolveRendererDom(): RendererDom {
  return {
    subtitleRoot: getRequiredElement<HTMLElement>('subtitleRoot'),
    subtitleContainer: getRequiredElement<HTMLElement>('subtitleContainer'),
    overlay: getRequiredElement<HTMLElement>('overlay'),
    overlayNotificationStack: getRequiredElement<HTMLDivElement>('overlayNotificationStack'),
    overlayNotificationHistory: getRequiredElement<HTMLElement>('overlayNotificationHistory'),
    controllerStatusToast: getRequiredElement<HTMLDivElement>('controllerStatusToast'),
    overlayErrorToast: getRequiredElement<HTMLDivElement>('overlayErrorToast'),
    secondarySubContainer: getRequiredElement<HTMLElement>('secondarySubContainer'),
    secondarySubRoot: getRequiredElement<HTMLElement>('secondarySubRoot'),

    jimakuModal: getRequiredElement<HTMLDivElement>('jimakuModal'),
    jimakuTitleInput: getRequiredElement<HTMLInputElement>('jimakuTitle'),
    jimakuSeasonInput: getRequiredElement<HTMLInputElement>('jimakuSeason'),
    jimakuEpisodeInput: getRequiredElement<HTMLInputElement>('jimakuEpisode'),
    jimakuSearchButton: getRequiredElement<HTMLButtonElement>('jimakuSearch'),
    jimakuCloseButton: getRequiredElement<HTMLButtonElement>('jimakuClose'),
    jimakuStatus: getRequiredElement<HTMLDivElement>('jimakuStatus'),
    jimakuEntriesSection: getRequiredElement<HTMLDivElement>('jimakuEntriesSection'),
    jimakuEntriesList: getRequiredElement<HTMLUListElement>('jimakuEntries'),
    jimakuFilesSection: getRequiredElement<HTMLDivElement>('jimakuFilesSection'),
    jimakuFilesList: getRequiredElement<HTMLUListElement>('jimakuFiles'),
    jimakuBroadenButton: getRequiredElement<HTMLButtonElement>('jimakuBroaden'),

    tsukihimeModal: getRequiredElement<HTMLDivElement>('tsukihimeModal'),
    tsukihimeTitleInput: getRequiredElement<HTMLInputElement>('tsukihimeTitle'),
    tsukihimeSeasonInput: getRequiredElement<HTMLInputElement>('tsukihimeSeason'),
    tsukihimeEpisodeInput: getRequiredElement<HTMLInputElement>('tsukihimeEpisode'),
    tsukihimeSearchButton: getRequiredElement<HTMLButtonElement>('tsukihimeSearch'),
    tsukihimeCloseButton: getRequiredElement<HTMLButtonElement>('tsukihimeClose'),
    tsukihimeTabSecondaryButton: getRequiredElement<HTMLButtonElement>('tsukihimeTabSecondary'),
    tsukihimeTabPrimaryButton: getRequiredElement<HTMLButtonElement>('tsukihimeTabPrimary'),
    tsukihimeStatus: getRequiredElement<HTMLDivElement>('tsukihimeStatus'),
    tsukihimeEntriesSection: getRequiredElement<HTMLDivElement>('tsukihimeEntriesSection'),
    tsukihimeEntriesList: getRequiredElement<HTMLUListElement>('tsukihimeEntries'),
    tsukihimeFilesSection: getRequiredElement<HTMLDivElement>('tsukihimeFilesSection'),
    tsukihimeFilesList: getRequiredElement<HTMLUListElement>('tsukihimeFiles'),

    youtubePickerModal: getRequiredElement<HTMLDivElement>('youtubePickerModal'),
    youtubePickerTitle: getRequiredElement<HTMLDivElement>('youtubePickerTitle'),
    youtubePickerPrimarySelect: getRequiredElement<HTMLSelectElement>('youtubePickerPrimarySelect'),
    youtubePickerSecondarySelect: getRequiredElement<HTMLSelectElement>(
      'youtubePickerSecondarySelect',
    ),
    youtubePickerContinueButton: getRequiredElement<HTMLButtonElement>(
      'youtubePickerContinueButton',
    ),
    youtubePickerCloseButton: getRequiredElement<HTMLButtonElement>('youtubePickerCloseButton'),
    youtubePickerStatus: getRequiredElement<HTMLDivElement>('youtubePickerStatus'),
    youtubePickerTracks: getRequiredElement<HTMLUListElement>('youtubePickerTracks'),

    kikuModal: getRequiredElement<HTMLDivElement>('kikuFieldGroupingModal'),
    kikuCard1: getRequiredElement<HTMLDivElement>('kikuCard1'),
    kikuCard2: getRequiredElement<HTMLDivElement>('kikuCard2'),
    kikuCard1Expression: getRequiredElement<HTMLDivElement>('kikuCard1Expression'),
    kikuCard2Expression: getRequiredElement<HTMLDivElement>('kikuCard2Expression'),
    kikuCard1Sentence: getRequiredElement<HTMLDivElement>('kikuCard1Sentence'),
    kikuCard2Sentence: getRequiredElement<HTMLDivElement>('kikuCard2Sentence'),
    kikuCard1Meta: getRequiredElement<HTMLDivElement>('kikuCard1Meta'),
    kikuCard2Meta: getRequiredElement<HTMLDivElement>('kikuCard2Meta'),
    kikuConfirmButton: getRequiredElement<HTMLButtonElement>('kikuConfirmButton'),
    kikuCancelButton: getRequiredElement<HTMLButtonElement>('kikuCancelButton'),
    kikuDeleteDuplicateCheckbox: getRequiredElement<HTMLInputElement>('kikuDeleteDuplicate'),
    kikuSelectionStep: getRequiredElement<HTMLDivElement>('kikuSelectionStep'),
    kikuPreviewStep: getRequiredElement<HTMLDivElement>('kikuPreviewStep'),
    kikuPreviewJson: getRequiredElement<HTMLPreElement>('kikuPreviewJson'),
    kikuPreviewCompactButton: getRequiredElement<HTMLButtonElement>('kikuPreviewCompact'),
    kikuPreviewFullButton: getRequiredElement<HTMLButtonElement>('kikuPreviewFull'),
    kikuPreviewError: getRequiredElement<HTMLDivElement>('kikuPreviewError'),
    kikuBackButton: getRequiredElement<HTMLButtonElement>('kikuBackButton'),
    kikuFinalConfirmButton: getRequiredElement<HTMLButtonElement>('kikuFinalConfirmButton'),
    kikuFinalCancelButton: getRequiredElement<HTMLButtonElement>('kikuFinalCancelButton'),
    kikuHint: getRequiredElement<HTMLDivElement>('kikuHint'),

    runtimeOptionsModal: getRequiredElement<HTMLDivElement>('runtimeOptionsModal'),
    runtimeOptionsClose: getRequiredElement<HTMLButtonElement>('runtimeOptionsClose'),
    runtimeOptionsList: getRequiredElement<HTMLUListElement>('runtimeOptionsList'),
    runtimeOptionsStatus: getRequiredElement<HTMLDivElement>('runtimeOptionsStatus'),

    characterDictionaryModal: getRequiredElement<HTMLDivElement>('characterDictionaryModal'),
    characterDictionaryClose: getRequiredElement<HTMLButtonElement>('characterDictionaryClose'),
    characterDictionaryOverrideTab: getRequiredElement<HTMLButtonElement>(
      'characterDictionaryOverrideTab',
    ),
    characterDictionaryManageTab: getRequiredElement<HTMLButtonElement>(
      'characterDictionaryManageTab',
    ),
    characterDictionarySummary: getRequiredElement<HTMLDivElement>('characterDictionarySummary'),
    characterDictionarySearchPanel: getRequiredElement<HTMLDivElement>(
      'characterDictionarySearchPanel',
    ),
    characterDictionarySearchInput: getRequiredElement<HTMLInputElement>(
      'characterDictionarySearchInput',
    ),
    characterDictionarySearchButton: getRequiredElement<HTMLButtonElement>(
      'characterDictionarySearchButton',
    ),
    characterDictionaryCurrent: getRequiredElement<HTMLDivElement>('characterDictionaryCurrent'),
    characterDictionaryCandidates: getRequiredElement<HTMLUListElement>(
      'characterDictionaryCandidates',
    ),
    characterDictionaryManagerPanel: getRequiredElement<HTMLDivElement>(
      'characterDictionaryManagerPanel',
    ),
    characterDictionaryManagedEntries: getRequiredElement<HTMLUListElement>(
      'characterDictionaryManagedEntries',
    ),
    characterDictionaryStatus: getRequiredElement<HTMLDivElement>('characterDictionaryStatus'),

    subsyncModal: getRequiredElement<HTMLDivElement>('subsyncModal'),
    subsyncCloseButton: getRequiredElement<HTMLButtonElement>('subsyncClose'),
    subsyncEngineAlass: getRequiredElement<HTMLInputElement>('subsyncEngineAlass'),
    subsyncEngineFfsubsync: getRequiredElement<HTMLInputElement>('subsyncEngineFfsubsync'),
    subsyncReferenceLabel: getRequiredElement<HTMLLabelElement>('subsyncReferenceLabel'),
    subsyncReferenceSelect: getRequiredElement<HTMLSelectElement>('subsyncReferenceSelect'),
    subsyncTargetLabel: getRequiredElement<HTMLLabelElement>('subsyncTargetLabel'),
    subsyncTargetSelect: getRequiredElement<HTMLSelectElement>('subsyncTargetSelect'),
    subsyncRunButton: getRequiredElement<HTMLButtonElement>('subsyncRun'),
    subsyncStatus: getRequiredElement<HTMLDivElement>('subsyncStatus'),

    controllerSelectModal: getRequiredElement<HTMLDivElement>('controllerSelectModal'),
    controllerSelectClose: getRequiredElement<HTMLButtonElement>('controllerSelectClose'),
    controllerSelectPicker: getRequiredElement<HTMLSelectElement>('controllerSelectPicker'),
    controllerSelectSummary: getRequiredElement<HTMLDivElement>('controllerSelectSummary'),
    controllerSelectStatus: getRequiredElement<HTMLDivElement>('controllerSelectStatus'),
    controllerConfigList: getRequiredElement<HTMLDivElement>('controllerConfigList'),
    controllerSelectSave: getRequiredElement<HTMLButtonElement>('controllerSelectSave'),

    controllerDebugModal: getRequiredElement<HTMLDivElement>('controllerDebugModal'),
    controllerDebugClose: getRequiredElement<HTMLButtonElement>('controllerDebugClose'),
    controllerDebugCopy: getRequiredElement<HTMLButtonElement>('controllerDebugCopy'),
    controllerDebugToast: getRequiredElement<HTMLDivElement>('controllerDebugToast'),
    controllerDebugStatus: getRequiredElement<HTMLDivElement>('controllerDebugStatus'),
    controllerDebugSummary: getRequiredElement<HTMLDivElement>('controllerDebugSummary'),
    controllerDebugAxes: getRequiredElement<HTMLPreElement>('controllerDebugAxes'),
    controllerDebugButtons: getRequiredElement<HTMLPreElement>('controllerDebugButtons'),
    controllerDebugButtonIndices: getRequiredElement<HTMLPreElement>(
      'controllerDebugButtonIndices',
    ),
    subtitleSidebarModal: getRequiredElement<HTMLDivElement>('subtitleSidebarModal'),
    subtitleSidebarContent: getRequiredElement<HTMLDivElement>('subtitleSidebarContent'),
    subtitleSidebarClose: getRequiredElement<HTMLButtonElement>('subtitleSidebarClose'),
    subtitleSidebarStatus: getRequiredElement<HTMLDivElement>('subtitleSidebarStatus'),
    subtitleSidebarList: getRequiredElement<HTMLUListElement>('subtitleSidebarList'),

    changelogModal: getRequiredElement<HTMLDivElement>('changelogModal'),
    changelogClose: getRequiredElement<HTMLButtonElement>('changelogClose'),
    changelogRefresh: getRequiredElement<HTMLButtonElement>('changelogRefresh'),
    changelogInstalled: getRequiredElement<HTMLSpanElement>('changelogInstalled'),
    changelogSource: getRequiredElement<HTMLSpanElement>('changelogSource'),
    changelogWarning: getRequiredElement<HTMLDivElement>('changelogWarning'),
    changelogStatus: getRequiredElement<HTMLDivElement>('changelogStatus'),
    changelogList: getRequiredElement<HTMLDivElement>('changelogList'),
    sessionHelpModal: getRequiredElement<HTMLDivElement>('sessionHelpModal'),
    sessionHelpClose: getRequiredElement<HTMLButtonElement>('sessionHelpClose'),
    sessionHelpShortcut: getRequiredElement<HTMLDivElement>('sessionHelpShortcut'),
    sessionHelpWarning: getRequiredElement<HTMLDivElement>('sessionHelpWarning'),
    sessionHelpStatus: getRequiredElement<HTMLDivElement>('sessionHelpStatus'),
    sessionHelpFilter: getRequiredElement<HTMLInputElement>('sessionHelpFilter'),
    sessionHelpContent: getRequiredElement<HTMLDivElement>('sessionHelpContent'),

    playlistBrowserModal: getRequiredElement<HTMLDivElement>('playlistBrowserModal'),
    playlistBrowserTitle: getRequiredElement<HTMLDivElement>('playlistBrowserTitle'),
    playlistBrowserStatus: getRequiredElement<HTMLDivElement>('playlistBrowserStatus'),
    playlistBrowserDirectoryList: getRequiredElement<HTMLUListElement>(
      'playlistBrowserDirectoryList',
    ),
    playlistBrowserPlaylistList: getRequiredElement<HTMLUListElement>(
      'playlistBrowserPlaylistList',
    ),
    playlistBrowserClose: getRequiredElement<HTMLButtonElement>('playlistBrowserClose'),

    animeBrowserModal: getRequiredElement<HTMLDivElement>('animeBrowserModal'),
    animeBrowserClose: getRequiredElement<HTMLButtonElement>('animeBrowserClose'),
    animeBrowserFrame: getRequiredElement<HTMLIFrameElement>('animeBrowserFrame'),
  };
}
