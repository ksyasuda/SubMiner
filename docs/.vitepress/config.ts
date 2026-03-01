const repositoryName = process.env.GITHUB_REPOSITORY?.split('/')[1];
const base = process.env.GITHUB_ACTIONS && repositoryName ? `/${repositoryName}/` : '/';

export default {
  title: 'SubMiner Docs',
  description:
    'SubMiner: an MPV immersion-mining overlay with Yomitan and AnkiConnect integration.',
  base,
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
      { text: 'Troubleshooting', link: '/troubleshooting' },
    ],
    sidebar: [
      {
        text: 'Getting Started',
        items: [
          { text: 'Overview', link: '/' },
          { text: 'Installation', link: '/installation' },
          { text: 'Launcher Script', link: '/launcher-script' },
          { text: 'Usage', link: '/usage' },
          { text: 'Mining Workflow', link: '/mining-workflow' },
          // { text: 'Feature Demos', link: '/demos' },
        ],
      },
      {
        text: 'Reference',
        items: [
          { text: 'Configuration', link: '/configuration' },
          { text: 'Keyboard Shortcuts', link: '/shortcuts' },
          { text: 'Anki Integration', link: '/anki-integration' },
          { text: 'Jellyfin Integration', link: '/jellyfin-integration' },
          { text: 'Immersion Tracking', link: '/immersion-tracking' },
          { text: 'JLPT Vocabulary', link: '/jlpt-vocab-bundle' },
          { text: 'MPV Plugin', link: '/mpv-plugin' },
          { text: 'Troubleshooting', link: '/troubleshooting' },
        ],
      },
      {
        text: 'Development',
        items: [
          { text: 'Building & Testing', link: '/development' },
          { text: 'Architecture', link: '/architecture' },
          { text: 'IPC + Runtime Contracts', link: '/ipc-contracts' },
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
      pattern: 'https://github.com/ksyasuda/SubMiner/edit/main/docs/:path',
      text: 'Edit this page on GitHub',
    },
    outline: { level: [2, 3], label: 'On this page' },
    externalLinkIcon: true,
    docFooter: { prev: 'Previous', next: 'Next' },
    returnToTopLabel: 'Back to top',
    socialLinks: [{ icon: 'github', link: 'https://github.com/ksyasuda/SubMiner' }],
  },
};
