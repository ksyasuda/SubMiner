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
- `/main/` serves development docs from `main` and is marked `noindex,follow`.
- `/v/<version>/` serves stable release archives.
- Prerelease tags do not update the docs site.

Keep Cloudflare Git auto-deploy disabled. The production deploy is `.github/workflows/docs-pages.yml`, which uploads `.tmp/docs-versioned-site` with `--branch main` so tag-triggered runs update Production instead of creating preview deployments.
