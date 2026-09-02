# SubMiner Docs

In-repo VitePress documentation source for SubMiner.

Internal architecture/workflow source of truth lives in `docs/README.md` at the repo root. Keep `docs-site/` user-facing.

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
- Production branch: `main`
- Automatic production and preview deployments: disabled
- Custom domain: `docs.subminer.moe` attached to Production
- Deployment path: GitHub Actions direct upload with Wrangler

The public docs root is stable-only:

- `/` serves the latest stable release docs.
- `/main/` serves development docs from `main`.
- `/v/<version>/` serves stable release archives.
- Prerelease tags do not update the docs site.

Only `/` is indexable. `/main/` and every `/v/<version>/` page carries a self-referential canonical plus `noindex,follow`, and the generated `_headers` file repeats that as an `X-Robots-Tag`. They stay crawlable so their links still resolve, but ~30 archived copies of every page would otherwise consume the crawl budget the current docs need. Only the root build emits `sitemap.xml`, and its `<lastmod>` dates come from `git log` against the tracked checkout at the released tag, because the build renders from an untracked snapshot that VitePress cannot date itself.

Keep Cloudflare Git auto-deploy disabled. The production deploy is `.github/workflows/docs-pages.yml`, which uploads `.tmp/docs-versioned-site` with `--branch main` so tag-triggered runs update Production instead of creating preview deployments.
