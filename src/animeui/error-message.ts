export interface AnimeBrowserErrorMessage {
  title: string;
  explanation: string;
  guidance: string;
  details: string;
}

const OPERATIONS: Record<string, string> = {
  getDetailsAnime: 'Could not load anime details',
  getEpisodeList: 'Could not load episodes',
  getVideoList: 'Could not resolve the video',
  getSearchAnime: 'Could not search this source',
  getPopularAnime: 'Could not browse this source',
};

/** Error messages survive both Electron IPC and the embedded browser's transport. */
export function describeAnimeBrowserError(message: string): AnimeBrowserErrorMessage {
  const bridge = /Anime bridge (\w+) failed(?: \(\d+\))?[.:]/.exec(message);
  const operation = bridge?.[1];
  const title = (operation && OPERATIONS[operation]) || 'Request failed';
  const details = message;

  if (
    bridge &&
    /NoSuchMethodError|NoClassDefFoundError|AbstractMethodError|(?:java\.lang\.Object|void|boolean) eu\.kanade\.[\w.$]+\(/.test(
      message,
    )
  ) {
    return {
      title,
      explanation:
        'This extension needs functionality that the installed extension bridge does not provide.',
      guidance:
        'Check Extensions for bridge updates. If none are available, try another source and include the technical details when reporting the issue.',
      details,
    };
  }

  if (bridge && /lateinit property \w+ has not been initialized/.test(message)) {
    return {
      title,
      explanation:
        'The extension bridge could not read a required field from the extension’s data.',
      guidance:
        'Check Extensions for extension and bridge updates. If this continues, try another source and include the technical details when reporting the issue.',
      details,
    };
  }

  if (bridge) {
    return {
      title,
      explanation: 'The extension bridge could not complete this request.',
      guidance:
        'Try the action again. If it keeps failing, check for updates in Extensions or try another source.',
      details,
    };
  }

  if (/^mpv could not play this stream|^Playback did not start\./.test(message)) {
    return {
      title: 'Could not start playback',
      explanation: 'mpv could not start the selected stream.',
      guidance:
        'Try playing the episode again to request a fresh stream, or choose another source.',
      details,
    };
  }

  return { title, explanation: message, guidance: '', details: '' };
}
