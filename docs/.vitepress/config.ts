const repositoryName = process.env.GITHUB_REPOSITORY?.split('/')[1];
const base = process.env.GITHUB_ACTIONS && repositoryName ? `/${repositoryName}/` : '/';

export default {
  title: 'SubMiner Docs',
  description: 'All-in-one sentence mining overlay for MPV with AnkiConnect and dictionary integration',
  base,
  appearance: 'dark',
  cleanUrls: true,
  lastUpdated: true,
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
          { text: 'Usage', link: '/usage' },
          { text: 'Mining Workflow', link: '/mining-workflow' },
        ],
      },
      {
        text: 'Reference',
        items: [
          { text: 'Configuration', link: '/configuration' },
          { text: 'Anki Integration', link: '/anki-integration' },
          { text: 'MPV Plugin', link: '/mpv-plugin' },
          { text: 'Troubleshooting', link: '/troubleshooting' },
        ],
      },
      {
        text: 'Development',
        items: [
          { text: 'Building & Testing', link: '/development' },
          { text: 'Architecture', link: '/architecture' },
        ],
      },
    ],
    socialLinks: [{ icon: 'github', link: 'https://github.com/ksyasuda/SubMiner' }],
  },
};
