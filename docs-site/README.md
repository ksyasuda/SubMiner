# SubMiner docs

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
- `/versions` (root build only) lists every published version.
- Prerelease tags do not update the docs site.

Only `/` is indexable. `/main/` and every `/v/<version>/` page carries a self-referential canonical plus `noindex,follow`, repeated as an `X-Robots-Tag` header (by the generated `_headers` file for `/main/`, by the archive function for `/v/`). They stay crawlable so their links still resolve, but ~30 archived copies of every page would otherwise consume the crawl budget the current docs need. Only the root build emits `sitemap.xml`, and its `<lastmod>` dates come from `git log` against the tracked checkout at the released tag, because the build renders from an untracked snapshot that VitePress cannot date itself.

Keep Cloudflare Git auto-deploy disabled. The production deploy is `.github/workflows/docs-pages.yml`, which uploads `.tmp/docs-versioned-site` with `--branch main` so tag-triggered runs update Production instead of creating preview deployments.

### Stable archives in R2

`/v/<version>/` archives are not part of the Pages deployment. Each one is built once, uploaded to an R2 bucket under `v/<version>/`, and served by the Pages Function in `functions/v/[[path]].ts`. Each deploy builds only archives the bucket is missing (an archive counts as present once its `_archive.json` marker exists), plus the root and `/main/` builds. Archive nav links to `/versions` instead of listing releases, so a new tag never invalidates old archives.

To re-render archives on purpose (theme change, docs fix), run the `Docs Pages` workflow manually with `rebuild_archives` set to comma-separated tags or `all`.

One-time setup:

- R2 bucket for archives; its name goes in the `DOCS_ARCHIVE_R2_BUCKET` repository variable.
- R2 API token with Object Read & Write on that bucket; its S3 credentials go in the `DOCS_ARCHIVE_R2_ACCESS_KEY_ID` and `DOCS_ARCHIVE_R2_SECRET_ACCESS_KEY` repository secrets.
- Pages project: Settings > Bindings > R2 bucket, variable name `DOCS_ARCHIVES`, pointing at the same bucket.

The first deploy after setup builds and uploads every stable archive; later deploys only add new tags.
