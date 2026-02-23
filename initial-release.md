# Initial Release Git History Cleanup Plan

## Recommendation

Given the current `main` history shape (325 commits, heavy refactor churn, and many internal maintenance commits), the best pre-release cleanup path is:

1. Preserve the old `main` history in remote backup refs.
2. Build a **new orphan history** from the current project snapshot.
3. Recreate history as **7 curated commits** grouped by release-facing domains.
4. Verify tree equality with old `main` and replace `origin/main` once.

This avoids fragile `rebase -i --root` over 300+ commits while still keeping meaningful archaeology for future maintenance.

## Why this is best for this repo

- Current history is dominated by refactor/noise commits (`refactor`: 138, `other`: 91).
- Commit velocity is very high and bursty (multiple 30-45 commit days), so interactive cleanup is risky/time-expensive.
- Codebase is already mature and release-shaped (Electron app + launcher + plugin + docs + CI), so snapshot regrouping is cleaner.
- You are pre-release and private; a one-time force update is low risk if coordinated.

---

## Pre-cutover checklist

Run from repository root.

```bash
git checkout main
git pull --ff-only origin main
```

Coordinate freeze:

- Pause merges/pushes to `main`.
- Ask teammates to stop rebasing onto `main` during cutover.

---

## Copy/paste command sequence

### 1) Safety backups (local + remote)

```bash
set -euo pipefail

git checkout main
git pull --ff-only origin main

OLD_MAIN_COMMIT="$(git rev-parse main)"
STAMP="$(date -u +%Y%m%dT%H%M%SZ)"

echo "Old main commit: ${OLD_MAIN_COMMIT}"
echo "Cutover stamp: ${STAMP}"

git tag "pre-release-history-backup-${STAMP}" "${OLD_MAIN_COMMIT}"
git branch "archive/pre-release-history-${STAMP}" "${OLD_MAIN_COMMIT}"

git push origin "pre-release-history-backup-${STAMP}"
git push origin "archive/pre-release-history-${STAMP}"
```

### 2) Handle dirty workspace safely

If `git status --short` is not empty:

```bash
git stash push -u -m "initial-release-history-cleanup-${STAMP}"
```

### 3) Build new clean history on orphan branch

```bash
git checkout --orphan release/main-clean
git reset
```

Now create curated commits.

```bash
# Commit 1: repository scaffolding and build toolchain
git add .gitignore .gitmodules .npmrc .prettierignore .prettierrc.json AGENTS.md LICENSE Makefile package.json bun.lock tsconfig.json tsconfig.renderer.json build scripts .github
git commit -m "chore: bootstrap repository tooling and release automation"

# Commit 2: core desktop app runtime
git add src
git commit -m "feat(core): add Electron runtime, services, and app composition"

# Commit 3: launcher and mpv plugin integration
git add launcher plugin subminer
git commit -m "feat(integration): add launcher and mpv plugin workflows"

# Commit 4: bundled assets and vendor payloads
git add assets vendor
git commit -m "feat(assets): bundle runtime assets and vendor dependencies"

# Commit 5: test suites and verification harnesses
git add tests
git commit -m "test: add regression suites and harness coverage"

# Commit 6: documentation and examples
git add README.md docs config.example.jsonc
git commit -m "docs: add setup guides, architecture docs, and config examples"

# Commit 7: project/backlog metadata and remaining files (if any)
git add -A
if ! git diff --cached --quiet; then
  git commit -m "chore: add project management metadata and remaining repository files"
fi
```

### 4) Validate tree equivalence to old main snapshot

```bash
git diff --stat "${OLD_MAIN_COMMIT}^{tree}" HEAD^{tree}
git diff --name-status "${OLD_MAIN_COMMIT}^{tree}" HEAD^{tree}
```

Expected: no output. If output exists, stop and inspect before pushing.

### 5) Replace `origin/main`

```bash
git push --force-with-lease origin release/main-clean:main
```

### 6) Tag cleaned initial-release baseline

```bash
git fetch origin
git checkout main
git reset --hard origin/main
git tag "v0.1.0-initial-clean-history"
git push origin "v0.1.0-initial-clean-history"
```

### 7) Restore local stashed work (if stashed earlier)

```bash
git stash list
git stash pop
```

---

## Team resync message (send after cutover)

```text
main history was rewritten for initial release cleanup.

If you have no local work:
  git fetch origin
  git checkout main
  git reset --hard origin/main

If you have local work:
  git fetch origin
  git checkout -b backup/<name>-before-main-rewrite
  # then rebase/cherry-pick your feature commits onto new origin/main
```

---

## Notes

- This plan intentionally avoids `git rebase -i --root` to reduce conflict risk.
- `--force-with-lease` is mandatory safety rail.
- Keep the backup tag/branch until after public release stabilization.
