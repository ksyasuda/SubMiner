const repositoryName = process.env.GITHUB_REPOSITORY?.split('/')[1];
const base = process.env.GITHUB_ACTIONS && repositoryName ? `/${repositoryName}/` : '/';

export default {
  title: 'SubMiner',
  description: 'Documentation for SubMiner',
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
    logo: '/assets/SubMiner.png',
    nav: [
      { text: 'Docs', link: '/' },
      { text: 'Installation', link: '/installation' },
      { text: 'Usage', link: '/usage' },
      { text: 'Configuration', link: '/configuration' },
      { text: 'Development', link: '/development' },
      { text: 'Architecture', link: '/architecture' },
    ],
    sidebar: [
      {
        text: 'Getting Started',
        items: [
          { text: 'Overview', link: '/' },
          { text: 'Installation', link: '/installation' },
          { text: 'Usage', link: '/usage' },
        ],
      },
      {
        text: 'Reference',
        items: [
          { text: 'Configuration', link: '/configuration' },
          { text: 'Development', link: '/development' },
          { text: 'Architecture', link: '/architecture' },
        ],
      },
    ],
    socialLinks: [{ icon: 'github', link: 'https://github.com/ksyasuda/SubMiner' }],
  },
};
