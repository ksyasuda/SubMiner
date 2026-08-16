type: internal
area: docs

- Excluded the `/main/` and `/v/<version>/` docs trees from search indexing with a self-referential canonical, `noindex,follow`, and a matching `X-Robots-Tag` header, so crawlers spend their budget on the current docs instead of ~30 archived copies of every page.
- Restored `<lastmod>` dates in the docs sitemap, which were silently dropped because production builds render from an untracked release snapshot.
