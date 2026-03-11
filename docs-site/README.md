# SubMiner Docs

In-repo VitePress documentation source for SubMiner.

## Local development

```bash
bun --cwd docs-site install
bun run docs:dev
```

Build and preview:

```bash
bun run docs:build
bun run docs:preview
bun run docs:test
```

Direct package commands still work from `docs-site/` if you prefer:

```bash
cd docs-site
bun install
bun run docs:dev
```

## Cloudflare Pages

- Git repo: `ksyasuda/SubMiner`
- Root directory: `docs-site`
- Build command: `bun run docs:build`
- Build output directory: `.vitepress/dist`
- Build watch paths: `docs-site/**`
