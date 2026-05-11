---
id: TASK-342
title: Improve docs homepage indexability
status: Done
assignee:
  - codex
created_date: '2026-05-11 04:53'
updated_date: '2026-05-11 05:06'
labels:
  - docs
  - seo
dependencies: []
references:
  - 'https://docs.subminer.moe/'
  - >-
    https://developers.google.com/search/docs/crawling-indexing/consolidate-duplicate-urls
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Google Search Console reports https://docs.subminer.moe/ as Crawled - currently not indexed. Investigate and improve the docs-site homepage and crawl signals so the root docs page has clear unique content, self-consistent canonical metadata, and no avoidable duplicate sitemap entry from VitePress README output. Keep the change scoped to docs-site SEO/content and verify with docs tests/build plus live-style generated HTML inspection.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Generated docs HTML for https://docs.subminer.moe/ declares a self-referential canonical URL and does not declare noindex.
- [x] #2 Generated sitemap does not advertise a duplicate /README docs page when the docs homepage is the intended canonical root.
- [x] #3 Docs tests cover the SEO/indexing signals so canonical and duplicate sitemap regressions are caught.
- [x] #4 Docs build/test commands pass or any blocker is documented with exact failure output.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing docs-site tests for canonical head generation and sitemap duplicate filtering. 2. Update VitePress config to emit canonical URLs and exclude /README from the sitemap. 3. Strengthen docs homepage content with unique overview/decision-path copy. 4. Run docs tests/build and inspect generated HTML/sitemap.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-05-11: Added docs SEO tests, root canonical generation, sitemap README filtering, and stronger docs homepage orientation copy. Verified generated index.html contains self canonical/no noindex and generated sitemap omits /README.

2026-05-11: Added changes/342-docs-indexability.md and verified changelog lint. Final verification: bun run --cwd docs-site test, bun run docs:build, bun x prettier --check touched docs/changelog files, bun run changelog:lint.

2026-05-11: User requested removal of the added homepage orientation block. Removed it, removed the matching content test, and updated the changelog fragment. Generated index.html no longer contains the removed headings while preserving canonical metadata.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Improved docs homepage indexability signals for https://docs.subminer.moe/. Added self-referential canonical generation for VitePress pages and filtered the duplicate /README page out of the generated sitemap. Added docs SEO regression coverage, kept the existing Plausible hostname test aligned with the shared hostname constant, and added a docs changelog fragment. Per follow-up request, the extra homepage orientation copy was removed before handoff. Verification passed: docs tests, docs build, targeted generated HTML/sitemap inspection, Prettier check for touched files, and changelog lint.
<!-- SECTION:FINAL_SUMMARY:END -->
