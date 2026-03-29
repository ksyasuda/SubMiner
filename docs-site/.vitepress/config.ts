export default {
  title: 'SubMiner Docs',
  description:
    'SubMiner: an MPV immersion-mining overlay with Yomitan and AnkiConnect integration.',
  head: [
    ['link', { rel: 'icon', href: '/favicon.ico', sizes: 'any' }],
    [
      'link',
      {
        rel: 'icon',
        type: 'image/png',
        href: '/favicon-32x32.png',
        sizes: '32x32',
      },
    ],
    [
      'link',
      {
        rel: 'icon',
        type: 'image/png',
        href: '/favicon-16x16.png',
        sizes: '16x16',
      },
    ],
    [
      'link',
      {
        rel: 'apple-touch-icon',
        href: '/apple-touch-icon.png',
        sizes: '180x180',
      },
    ],
  ],
  appearance: 'dark',
  cleanUrls: true,
  metaChunk: true,
  sitemap: { hostname: 'https://docs.subminer.moe' },
  lastUpdated: true,
  srcExclude: ['subagents/**'],
  markdown: {
    theme: {
      light: 'catppuccin-latte',
      dark: 'catppuccin-macchiato',
    },
  },
  themeConfig: {
    logo: {
      light: '/assets/SubMiner.png',
      dark: '/assets/SubMiner.png',
    },
    siteTitle: 'SubMiner Docs',
    nav: [
      { text: 'Home', link: '/' },
      { text: 'Get Started', link: '/installation' },
      { text: 'Mining', link: '/mining-workflow' },
      { text: 'Configuration', link: '/configuration' },
      { text: 'Changelog', link: '/changelog' },
      { text: 'Troubleshooting', link: '/troubleshooting' },
    ],
    sidebar: [
      {
        text: 'Getting Started',
        items: [
          { text: 'Overview', link: '/' },
          { text: 'Installation', link: '/installation' },
          { text: 'Usage', link: '/usage' },
          { text: 'Mining Workflow', link: '/mining-workflow' },
          { text: 'Launcher Script', link: '/launcher-script' },
        ],
      },
      {
        text: 'Reference',
        items: [
          { text: 'Configuration', link: '/configuration' },
          { text: 'Keyboard Shortcuts', link: '/shortcuts' },
          { text: 'Subtitle Annotations', link: '/subtitle-annotations' },
          { text: 'Subtitle Sidebar', link: '/subtitle-sidebar' },
          { text: 'Immersion Tracking', link: '/immersion-tracking' },
          { text: 'JLPT Vocabulary Bundle', link: '/jlpt-vocab-bundle' },
          { text: 'Troubleshooting', link: '/troubleshooting' },
        ],
      },
      {
        text: 'Integrations',
        items: [
          { text: 'MPV Plugin', link: '/mpv-plugin' },
          { text: 'Anki', link: '/anki-integration' },
          { text: 'Jellyfin', link: '/jellyfin-integration' },
          { text: 'YouTube', link: '/youtube-integration' },
          { text: 'Jimaku', link: '/jimaku-integration' },
          { text: 'AniList', link: '/anilist-integration' },
          { text: 'Character Dictionary', link: '/character-dictionary' },
        ],
      },
      {
        text: 'Development',
        items: [
          { text: 'Building & Testing', link: '/development' },
          { text: 'Architecture', link: '/architecture' },
          { text: 'IPC + Runtime Contracts', link: '/ipc-contracts' },
          { text: 'WebSocket + Texthooker API', link: '/websocket-texthooker-api' },
          { text: 'Changelog', link: '/changelog' },
        ],
      },
    ],
    search: {
      provider: 'local',
    },
    footer: {
      message: 'Released under the GPL-3.0 License.',
      copyright: 'Copyright © 2026-present sudacode',
    },
    editLink: {
      pattern: 'https://github.com/ksyasuda/SubMiner/edit/main/docs-site/:path',
      text: 'Edit this page on GitHub',
    },
    outline: { level: [2, 3], label: 'On this page' },
    externalLinkIcon: true,
    docFooter: { prev: 'Previous', next: 'Next' },
    returnToTopLabel: 'Back to top',
    socialLinks: [{ icon: 'github', link: 'https://github.com/ksyasuda/SubMiner' }],
  },
};
