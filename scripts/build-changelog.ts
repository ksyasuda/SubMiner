import * as fs from 'node:fs';
import * as path from 'node:path';
import { execFileSync } from 'node:child_process';

type RunClaude = (input: string, args: string[]) => string;

// A single PR's contribution, resolved from the fragment files released in this
// cycle. Used to append GitHub-style attribution to the release notes.
type Contribution = {
  prNumber: number;
  login: string;
  title: string;
  isFirstContribution: boolean;
};

// Resolves the contributions behind a set of changelog fragment paths. Injected
// in tests so we never hit git/gh; the default implementation walks git history
// and the GitHub API.
type ResolveContributions = (fragmentPaths: string[], cwd: string) => Contribution[];

// One changelog fragment's change between the previous prerelease tag and the
// working tree. `before` is the content at the tag, `after` the current content.
export type FragmentDeltaEntry = {
  path: string;
  status: 'added' | 'modified' | 'deleted';
  before?: string;
  after?: string;
};

type ChangelogFsDeps = {
  existsSync?: (candidate: string) => boolean;
  mkdirSync?: (candidate: string, options: { recursive: true }) => void;
  readFileSync?: (candidate: string, encoding: BufferEncoding) => string;
  readdirSync?: (candidate: string, options: { withFileTypes: true }) => fs.Dirent[];
  rmSync?: (candidate: string) => void;
  writeFileSync?: (candidate: string, content: string, encoding: BufferEncoding) => void;
  log?: (message: string) => void;
  runClaude?: RunClaude;
  resolveContributions?: ResolveContributions;
  listPrereleaseTags?: (cwd: string, baseVersion: string) => string[];
  resolveFragmentDelta?: (cwd: string, previousTag: string) => FragmentDeltaEntry[];
};

type PolishMode = 'changelog' | 'release-notes';

type ChangelogOptions = {
  cwd?: string;
  date?: string;
  version?: string;
  deps?: ChangelogFsDeps;
};

type FragmentType = 'added' | 'changed' | 'fixed' | 'docs' | 'internal';

type ChangeFragment = {
  area: string;
  breaking: boolean;
  bullets: string[];
  path: string;
  type: FragmentType;
};

type PullRequestChangelogOptions = {
  changedEntries: Array<{
    path: string;
    status: string;
  }>;
  changedLabels?: string[];
};

const RELEASE_NOTES_PATH = path.join('release', 'release-notes.md');
const PRERELEASE_NOTES_PATH = path.join('release', 'prerelease-notes.md');
const CHANGELOG_HEADER = '# Changelog';
const CHANGE_TYPES: FragmentType[] = ['added', 'changed', 'fixed', 'docs', 'internal'];
const SKIP_CHANGELOG_LABEL = 'skip-changelog';

function normalizeVersion(version: string): string {
  return version.replace(/^v/, '');
}

function resolveDate(date?: string): string {
  return date ?? new Date().toISOString().slice(0, 10);
}

function resolvePackageVersion(
  cwd: string,
  readFileSync: (candidate: string, encoding: BufferEncoding) => string,
): string {
  const packageJsonPath = path.join(cwd, 'package.json');
  const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8')) as { version?: string };
  if (!packageJson.version) {
    throw new Error(`Missing package.json version at ${packageJsonPath}`);
  }
  return normalizeVersion(packageJson.version);
}

function resolveVersion(options: Pick<ChangelogOptions, 'cwd' | 'version' | 'deps'>): string {
  const cwd = options.cwd ?? process.cwd();
  const readFileSync = options.deps?.readFileSync ?? fs.readFileSync;
  return normalizeVersion(options.version ?? resolvePackageVersion(cwd, readFileSync));
}

function isSupportedPrereleaseVersion(version: string): boolean {
  return /^\d+\.\d+\.\d+-(beta|rc)\.\d+$/u.test(normalizeVersion(version));
}

function resolvePrereleaseBaseVersion(version: string): string {
  const match = /^(\d+\.\d+\.\d+)-(?:beta|rc)\.\d+$/u.exec(normalizeVersion(version));
  if (!match) {
    throw new Error(
      `Unsupported prerelease version (${version}). Expected x.y.z-beta.N or x.y.z-rc.N.`,
    );
  }
  return match[1]!;
}

// The marker records which exact prerelease the committed notes were generated
// for (and which prior tag the delta section compares against), so CI can
// reject notes that were prepared for a different beta/RC.
function renderPrereleaseVersionMarker(version: string, previousTag: string | null): string {
  const since = previousTag ? `; since: ${previousTag}` : '';
  return `<!-- prerelease-version: ${normalizeVersion(version)}${since} -->`;
}

export function extractPrereleaseVersionMarker(notes: string): string | null {
  return (
    /<!--\s*prerelease-version:\s*(\d+\.\d+\.\d+-(?:beta|rc)\.\d+)(?:;\s*since:\s*\S+)?\s*-->/u.exec(
      notes,
    )?.[1] ?? null
  );
}

// Legacy marker written before the per-version marker existed. Still accepted
// when deciding whether existing notes can seed the cumulative baseline.
function extractPrereleaseBaseVersionMarker(notes: string): string | null {
  const fullVersion = extractPrereleaseVersionMarker(notes);
  if (fullVersion) {
    return resolvePrereleaseBaseVersion(fullVersion);
  }
  return /<!--\s*prerelease-base-version:\s*(\d+\.\d+\.\d+)\s*-->/u.exec(notes)?.[1] ?? null;
}

const DELTA_SECTION_HEADING_PREFIX = '## Changes since ';

// Removes the previous run's "Changes since" section so the cumulative baseline
// fed back to Claude never carries a stale beta-to-beta delta.
function stripDeltaSection(notes: string): string {
  const lines = notes.split(/\r?\n/);
  const start = lines.findIndex((line) => line.startsWith(DELTA_SECTION_HEADING_PREFIX));
  if (start === -1) {
    return notes;
  }
  let end = lines.length;
  for (let index = start + 1; index < lines.length; index += 1) {
    if (lines[index]!.startsWith('## ')) {
      end = index;
      break;
    }
  }
  return [...lines.slice(0, start), ...lines.slice(end)].join('\n');
}

function stripPrereleaseMetadata(notes: string): string {
  return notes
    .replace(/<!--\s*prerelease-version:[^>]*-->\s*/u, '')
    .replace(/<!--\s*prerelease-base-version:\s*\d+\.\d+\.\d+\s*-->\s*/u, '')
    .trim();
}

function resolveReusablePrereleaseNotes(notes: string, version: string): string | undefined {
  const existingBaseVersion = extractPrereleaseBaseVersionMarker(notes);
  if (existingBaseVersion !== resolvePrereleaseBaseVersion(version)) {
    return undefined;
  }
  return stripPrereleaseMetadata(stripDeltaSection(notes));
}

type ParsedPrereleaseTag = {
  tag: string;
  base: string;
  channel: 'beta' | 'rc';
  iteration: number;
};

function parsePrereleaseTag(tag: string): ParsedPrereleaseTag | null {
  const match = /^v?(\d+\.\d+\.\d+)-(beta|rc)\.(\d+)$/u.exec(tag.trim());
  if (!match) {
    return null;
  }
  return {
    tag: tag.trim(),
    base: match[1]!,
    channel: match[2] as 'beta' | 'rc',
    iteration: Number.parseInt(match[3]!, 10),
  };
}

// Semver prerelease order: every beta sorts before every rc, then numerically.
function comparePrereleaseTags(a: ParsedPrereleaseTag, b: ParsedPrereleaseTag): number {
  if (a.channel !== b.channel) {
    return a.channel === 'beta' ? -1 : 1;
  }
  return a.iteration - b.iteration;
}

// Picks the newest prerelease tag for the same base version that strictly
// precedes the version being released. Returns null for the first prerelease.
export function selectPreviousPrereleaseTag(tags: string[], version: string): string | null {
  const current = parsePrereleaseTag(normalizeVersion(version));
  if (!current) {
    return null;
  }

  const candidates = tags
    .map(parsePrereleaseTag)
    .filter((parsed): parsed is ParsedPrereleaseTag => parsed !== null)
    .filter((parsed) => parsed.base === current.base)
    .filter((parsed) => comparePrereleaseTags(parsed, current) < 0)
    .sort(comparePrereleaseTags);

  return candidates[candidates.length - 1]?.tag ?? null;
}

function defaultListPrereleaseTags(cwd: string, baseVersion: string): string[] {
  return execFileSync('git', ['tag', '--list', `v${baseVersion}-beta.*`, `v${baseVersion}-rc.*`], {
    cwd,
    encoding: 'utf8',
  })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

// Diffs changes/*.md between the previous prerelease tag and the working tree.
// Renamed fragments are treated as modifications of the new path.
//
// Like every other path in this script, git paths are resolved against `cwd`,
// which is the project root and also the repository root. Callers that point
// `cwd` elsewhere already fail earlier and loudly, when package.json and
// changes/ come back missing.
function defaultResolveFragmentDelta(cwd: string, previousTag: string): FragmentDeltaEntry[] {
  const output = execFileSync(
    'git',
    ['diff', '--name-status', '--find-renames', previousTag, '--', 'changes'],
    { cwd, encoding: 'utf8' },
  );
  const showAtTag = (fragmentPath: string): string =>
    execFileSync('git', ['show', `${previousTag}:${fragmentPath}`], { cwd, encoding: 'utf8' });
  const readCurrent = (fragmentPath: string): string =>
    fs.readFileSync(path.join(cwd, fragmentPath), 'utf8');

  const entries: FragmentDeltaEntry[] = [];
  for (const line of output.split(/\r?\n/)) {
    if (!line.trim()) {
      continue;
    }
    const [status = '', ...paths] = line.split('\t');
    const oldPath = paths[0] ?? '';
    const newPath = paths[paths.length - 1] ?? '';
    if (!isFragmentPath(newPath) && !isFragmentPath(oldPath)) {
      continue;
    }

    if (status.startsWith('A')) {
      entries.push({ path: newPath, status: 'added', after: readCurrent(newPath) });
    } else if (status.startsWith('D')) {
      entries.push({ path: oldPath, status: 'deleted', before: showAtTag(oldPath) });
    } else if (status.startsWith('M') || status.startsWith('R')) {
      entries.push({
        path: newPath,
        status: 'modified',
        before: showAtTag(oldPath),
        after: readCurrent(newPath),
      });
    }
  }

  // git diff misses fragments that exist only in the working tree; treat
  // untracked fragments as additions so a pre-commit run still sees them.
  const untracked = execFileSync(
    'git',
    ['ls-files', '--others', '--exclude-standard', '--', 'changes'],
    { cwd, encoding: 'utf8' },
  )
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((candidate) => candidate && isFragmentPath(candidate));
  for (const fragmentPath of untracked) {
    entries.push({ path: fragmentPath, status: 'added', after: readCurrent(fragmentPath) });
  }

  return entries;
}

function verifyRequestedVersionMatchesPackageVersion(
  options: Pick<ChangelogOptions, 'cwd' | 'version' | 'deps'>,
): void {
  if (!options.version) {
    return;
  }

  const cwd = options.cwd ?? process.cwd();
  const existsSync = options.deps?.existsSync ?? fs.existsSync;
  const readFileSync = options.deps?.readFileSync ?? fs.readFileSync;
  const packageJsonPath = path.join(cwd, 'package.json');
  if (!existsSync(packageJsonPath)) {
    return;
  }

  const packageVersion = resolvePackageVersion(cwd, readFileSync);
  const requestedVersion = normalizeVersion(options.version);

  if (packageVersion !== requestedVersion) {
    throw new Error(
      `package.json version (${packageVersion}) does not match requested release version (${requestedVersion}).`,
    );
  }
}

function resolveChangesDir(cwd: string): string {
  return path.join(cwd, 'changes');
}

function resolveFragmentPaths(cwd: string, deps?: ChangelogFsDeps): string[] {
  const changesDir = resolveChangesDir(cwd);
  const existsSync = deps?.existsSync ?? fs.existsSync;
  const readdirSync = deps?.readdirSync ?? fs.readdirSync;

  if (!existsSync(changesDir)) {
    return [];
  }

  return readdirSync(changesDir, { withFileTypes: true })
    .filter(
      (entry) =>
        entry.isFile() && entry.name.endsWith('.md') && entry.name.toLowerCase() !== 'readme.md',
    )
    .map((entry) => path.join(changesDir, entry.name))
    .sort();
}

function normalizeFragmentBullets(content: string): string[] {
  const lines = content
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const match = /^[-*]\s+(.*)$/.exec(line);
      return `- ${(match?.[1] ?? line).trim()}`;
    });

  if (lines.length === 0) {
    throw new Error('Changelog fragment cannot be empty.');
  }

  return lines;
}

function parseFragmentMetadata(
  content: string,
  fragmentPath: string,
): {
  area: string;
  body: string;
  breaking: boolean;
  type: FragmentType;
} {
  const lines = content.split(/\r?\n/);
  let index = 0;

  while (index < lines.length && !(lines[index] ?? '').trim()) {
    index += 1;
  }

  const metadata = new Map<string, string>();
  while (index < lines.length) {
    const trimmed = (lines[index] ?? '').trim();
    if (!trimmed) {
      index += 1;
      break;
    }

    const match = /^([a-z]+):\s*(.+)$/.exec(trimmed);
    if (!match) {
      break;
    }

    const [, rawKey = '', rawValue = ''] = match;
    metadata.set(rawKey, rawValue.trim());
    index += 1;
  }

  const type = metadata.get('type');
  if (!type || !CHANGE_TYPES.includes(type as FragmentType)) {
    throw new Error(`${fragmentPath} must declare type as one of: ${CHANGE_TYPES.join(', ')}.`);
  }

  const area = metadata.get('area');
  if (!area) {
    throw new Error(`${fragmentPath} must declare area.`);
  }

  const body = lines.slice(index).join('\n').trim();
  if (!body) {
    throw new Error(`${fragmentPath} must include at least one changelog bullet.`);
  }

  const breaking = metadata.get('breaking')?.toLowerCase() === 'true';

  return {
    area,
    body,
    breaking,
    type: type as FragmentType,
  };
}

function readChangeFragments(cwd: string, deps?: ChangelogFsDeps): ChangeFragment[] {
  const readFileSync = deps?.readFileSync ?? fs.readFileSync;
  return resolveFragmentPaths(cwd, deps).map((fragmentPath) => {
    const parsed = parseFragmentMetadata(readFileSync(fragmentPath, 'utf8'), fragmentPath);
    return {
      area: parsed.area,
      breaking: parsed.breaking,
      bullets: normalizeFragmentBullets(parsed.body),
      path: fragmentPath,
      type: parsed.type,
    };
  });
}

// We deliberately don't pass --bare here. --bare skips OAuth/keychain reads and
// requires ANTHROPIC_API_KEY, which most Claude Code users don't have set up.
// The polish prompt is self-contained and doesn't need tools, so loading the
// user's hooks/MCP/CLAUDE.md is harmless overhead.
const CLAUDE_CLI_ARGS = [
  '-p',
  '--model',
  'sonnet',
  '--permission-mode',
  'bypassPermissions',
  '--output-format',
  'text',
];

const SECTION_HEADER_PATTERN = /^### (Breaking Changes|Added|Changed|Fixed|Docs|Internal)$/m;

const POLISH_PROMPT_INSTRUCTIONS = `You are formatting a software release changelog for end users of SubMiner, an Electron app for Japanese sentence mining.

You will receive a list of FRAGMENT entries below. Each fragment has metadata (type, area, breaking) and one or more bullet points written by the engineer who shipped that change. Your job is to merge, dedupe, and rewrite these fragments into a polished, user-facing release body.

## Release Outcome Rules

- Treat the fragment list as one cumulative release outcome, not a chronological log of beta/RC churn.
- Put a fragment in ### Breaking Changes only if the final release requires action from users upgrading from the previous stable release. A breaking: true marker is a warning to preserve and evaluate the substance, not an automatic section choice.
- If a breaking or fixed fragment only changes behavior introduced by another pending fragment in the same release cycle, merge it into the final Added or Changed bullet. Example: if fragments first add a Config window and later rename or fix it as a Settings window, output one Settings Window bullet under Added, not separate Config window, Breaking Changes, or Fixed bullets.
- Multiple fixes within the same prerelease cycle should collapse into one current-state bullet that describes the final behavior.

## Output Rules

1. Output Markdown ONLY. No preamble, no commentary, no closing remarks. Start directly with the first section heading.
2. Use these section headings, in this order, omitting any that have no bullets:
   ### Breaking Changes
   ### Added
   ### Changed
   ### Fixed
   ### Docs
3. In MODE: changelog only, append a final section after Docs:
   <details>
   <summary>Internal changes</summary>

   ### Internal
   - …

   </details>
   Do not include the Internal section at all in MODE: release-notes; internal fragments will not be present in the input for that mode.
4. Each top-level change item should:
   - Lead with a short feature/area name in title case. Pick the name from the fragment's bullet content, not the raw 'area:' slug.
   - Be written in user-facing language. Drop implementation jargon, internal class names, file paths, and PR numbers.
   - Be merged with related bullets when possible. If five fragments all touch Windows overlay z-order/focus/restore, write one or two bullets that summarize the overall improvement instead of five.
   - Drop bullets that only describe PR housekeeping, CodeRabbit follow-ups, or test-only changes that don't affect users.
   - Preserve the substance of breaking changes that remain breaking after applying the Release Outcome Rules. Do not soften or omit them.
5. In both modes, split every item into one nested bullet per distinct change. Write a short bold name on the top-level bullet, then indent the details two spaces:
   - **Playlist Browser**:
     - Saved shows now open without rescanning the library.
     - The picker remembers the last folder you browsed between launches.
   Each nested bullet covers exactly one change, behavior, or user-visible outcome. Never stack several distinct changes into one long paragraph-shaped bullet.
   Aim for two to five nested bullets per item. When an item genuinely has only one thing to say, put it inline on the top-level bullet ("- **Playlist Browser**: Saved shows now open without rescanning the library.") instead of emitting a single nested bullet.
   Keep nested bullets short, concrete, and readable by non-technical users. Avoid paragraph-style release-note bullets.
   Bullets inside the Internal section may stay single-level.
6. In MODE: release-notes, nested bullets should also cover user benefit and any user action or compatibility note when useful. Do not require the exact nested labels; natural phrasing is fine. Omit the action bullet when no action is needed.
7. Do not invent features. Every bullet must be grounded in the input fragments.
8. Do not include the version heading (## v...) — that wrapper is added by the caller.

The input begins below.

`;

function defaultRunClaude(input: string, args: string[]): string {
  try {
    return execFileSync('claude', args, {
      input,
      encoding: 'utf8',
      maxBuffer: 10 * 1024 * 1024,
      stdio: ['pipe', 'pipe', 'inherit'],
    });
  } catch (error) {
    const err = error as NodeJS.ErrnoException;
    if (err.code === 'ENOENT') {
      throw new Error(
        "claude CLI not found on PATH. Install Claude Code (https://claude.com/claude-code) and ensure 'claude' is on your PATH before running changelog:build.",
      );
    }
    throw new Error(`claude CLI invocation failed: ${err.message}`);
  }
}

function resolveFragmentRelativePath(fragmentPath: string, cwd: string): string {
  return path.relative(cwd, fragmentPath).split(path.sep).join('/');
}

export function shouldSkipDefaultContributionLookup(
  env: Partial<Record<'GITHUB_ACTIONS' | 'GH_TOKEN' | 'GITHUB_TOKEN', string>> = process.env,
): boolean {
  return env.GITHUB_ACTIONS === 'true' && !env.GH_TOKEN && !env.GITHUB_TOKEN;
}

// Walks git history + the GitHub API to attribute each released fragment to the
// PR (and author) that introduced it. One git call and one gh call per fragment,
// plus one gh call per unique author for the first-contribution check. Best
// effort: if gh is unavailable/unauthenticated or any lookup fails, we warn and
// drop attribution rather than failing the release.
function defaultResolveContributions(fragmentPaths: string[], cwd: string): Contribution[] {
  if (fragmentPaths.length === 0) {
    return [];
  }
  if (shouldSkipDefaultContributionLookup()) {
    return [];
  }

  try {
    const slug = execFileSync(
      'gh',
      ['repo', 'view', '--json', 'nameWithOwner', '--jq', '.nameWithOwner'],
      {
        cwd,
        encoding: 'utf8',
      },
    ).trim();
    if (!slug) {
      return [];
    }

    const byPr = new Map<number, Contribution>();
    for (const fragmentPath of fragmentPaths) {
      const relativePath = resolveFragmentRelativePath(fragmentPath, cwd);
      // git log lists newest first, so the commit that *added* the file is the
      // last line of the --diff-filter=A history.
      const addingSha = execFileSync(
        'git',
        ['log', '--diff-filter=A', '--follow', '--format=%H', '--', relativePath],
        { cwd, encoding: 'utf8' },
      )
        .trim()
        .split(/\r?\n/)
        .filter(Boolean)
        .pop();
      if (!addingSha) {
        continue;
      }

      const prRaw = execFileSync(
        'gh',
        [
          'api',
          `repos/${slug}/commits/${addingSha}/pulls`,
          '--jq',
          '.[0] // empty | {number, login: .user.login, title}',
        ],
        { cwd, encoding: 'utf8' },
      ).trim();
      if (!prRaw) {
        continue;
      }

      const pr = JSON.parse(prRaw) as { number?: number; login?: string; title?: string };
      if (typeof pr.number !== 'number' || !pr.login || !pr.title) {
        continue;
      }
      if (!byPr.has(pr.number)) {
        byPr.set(pr.number, {
          prNumber: pr.number,
          login: pr.login,
          title: pr.title,
          isFirstContribution: false,
        });
      }
    }

    const firstPrByAuthor = new Map<string, number | null>();
    for (const contribution of byPr.values()) {
      if (!firstPrByAuthor.has(contribution.login)) {
        const firstRaw = execFileSync(
          'gh',
          [
            'api',
            '-X',
            'GET',
            'search/issues',
            '-f',
            `q=repo:${slug} is:pr is:merged author:${contribution.login}`,
            '-f',
            'sort=created',
            '-f',
            'order=asc',
            '-f',
            'per_page=1',
            '--jq',
            '.items[0].number // empty',
          ],
          { cwd, encoding: 'utf8' },
        ).trim();
        firstPrByAuthor.set(contribution.login, firstRaw ? Number.parseInt(firstRaw, 10) : null);
      }
      const firstPr = firstPrByAuthor.get(contribution.login) ?? null;
      contribution.isFirstContribution = firstPr !== null && firstPr === contribution.prNumber;
    }

    return [...byPr.values()].sort((a, b) => a.prNumber - b.prNumber);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn(`Skipping contributor attribution: ${message}`);
    return [];
  }
}

function resolveContributionsForFragments(
  fragments: ChangeFragment[],
  cwd: string,
  deps?: ChangelogFsDeps,
): Contribution[] {
  const resolve = deps?.resolveContributions ?? defaultResolveContributions;
  return resolve(
    fragments.filter((fragment) => fragment.type !== 'internal').map((fragment) => fragment.path),
    cwd,
  );
}

function isWhatsChangedHeading(line: string): boolean {
  return line === "## What's Changed" || line === '## What’s Changed';
}

function extractContributorSections(releaseNotes: string): string[] {
  const lines = releaseNotes.split(/\r?\n/);
  const start = lines.findIndex(isWhatsChangedHeading);
  if (start === -1) {
    return [];
  }

  let end = lines.length;
  for (let index = start + 1; index < lines.length; index += 1) {
    const line = lines[index]!;
    if (line.startsWith('## ') && !isWhatsChangedHeading(line) && line !== '## New Contributors') {
      end = index;
      break;
    }
  }

  const block = lines.slice(start, end);
  while (block.length > 0 && block[block.length - 1] === '') {
    block.pop();
  }
  if (block.length === 0) {
    return [];
  }

  block[0] = "## What's Changed";
  block.push('');
  return block;
}

function renderContributorsSections(contributions: Contribution[]): string[] {
  if (contributions.length === 0) {
    return [];
  }

  const lines: string[] = ["## What's Changed", ''];
  for (const contribution of contributions) {
    lines.push(`- ${contribution.title} by @${contribution.login} in #${contribution.prNumber}`);
  }

  const firstTimers = contributions.filter((contribution) => contribution.isFirstContribution);
  if (firstTimers.length > 0) {
    lines.push('', '## New Contributors', '');
    for (const contribution of firstTimers) {
      lines.push(
        `- @${contribution.login} made their first contribution in #${contribution.prNumber}`,
      );
    }
  }

  lines.push('');
  return lines;
}

function serializeFragmentsForPrompt(
  fragments: ChangeFragment[],
  mode: PolishMode,
  version: string,
  date?: string,
  existingReleaseNotes?: string,
): string {
  const header: string[] = [`MODE: ${mode}`, `VERSION: ${version}`];
  if (date) {
    header.push(`DATE: ${date}`);
  }

  const fragmentBlocks = fragments.map((fragment) => {
    const relativePath = fragment.path.replace(/^.*?(changes\/.*)$/u, '$1');
    return [
      `FRAGMENT ${relativePath}`,
      `type: ${fragment.type}`,
      `area: ${fragment.area}`,
      `breaking: ${fragment.breaking}`,
      ...fragment.bullets,
    ].join('\n');
  });

  const existingNotesBlock = existingReleaseNotes?.trim()
    ? ['EXISTING PRERELEASE NOTES', existingReleaseNotes.trim()]
    : [];

  return [...header, '', ...existingNotesBlock, '', ...fragmentBlocks].join('\n\n');
}

function validatePolishedOutput(
  output: string,
  mode: PolishMode,
  hasInternalFragments: boolean,
): string {
  const trimmed = output.trim();
  if (!trimmed) {
    throw new Error('claude returned empty output for changelog polish.');
  }
  if (!SECTION_HEADER_PATTERN.test(trimmed)) {
    throw new Error(
      `claude output is missing the expected section heading (### Added/Changed/Fixed/Docs/Breaking Changes). Got:\n${trimmed.slice(0, 400)}`,
    );
  }
  if (mode === 'changelog' && hasInternalFragments) {
    if (!/<details>[\s\S]*<summary>[^<]*Internal[^<]*<\/summary>/m.test(trimmed)) {
      throw new Error(
        'claude output is missing the expected <details><summary>Internal changes</summary> wrapper for the Internal section.',
      );
    }
  }
  return trimmed;
}

function polishFragmentsWithClaude(
  fragments: ChangeFragment[],
  options: {
    mode: PolishMode;
    version: string;
    date?: string;
    existingReleaseNotes?: string;
    deps?: ChangelogFsDeps;
  },
): string {
  const { mode, version, date, existingReleaseNotes } = options;
  const runClaude = options.deps?.runClaude ?? defaultRunClaude;

  const filtered =
    mode === 'release-notes'
      ? fragments.filter((fragment) => fragment.type !== 'internal')
      : fragments;
  const hasInternalFragments =
    mode === 'changelog' && fragments.some((fragment) => fragment.type === 'internal');

  if (filtered.length === 0) {
    throw new Error(
      mode === 'release-notes'
        ? 'No user-facing changelog fragments found in changes/ (only internal fragments are present, which are dropped from release notes).'
        : 'No changelog fragments found in changes/.',
    );
  }

  const reuseInstructions = existingReleaseNotes?.trim()
    ? [
        '## Existing Prerelease Notes',
        '',
        'The input includes EXISTING PRERELEASE NOTES before the fragment list. Existing prerelease notes are a baseline, not an immutable changelog. Reuse reviewed highlight bullets when they still describe the current outcome, but replace stale beta or RC wording when new fragments supersede it. Merge in only new or changed fragment material, and deduplicate instead of restating existing bullets. Output only the final highlights body using the section headings above; do not include the prerelease disclaimer, any "Changes since" section, or the Installation or Assets sections.',
        '',
      ].join('\n')
    : '';
  const prompt =
    POLISH_PROMPT_INSTRUCTIONS +
    reuseInstructions +
    serializeFragmentsForPrompt(filtered, mode, version, date, existingReleaseNotes);
  const output = runClaude(prompt, CLAUDE_CLI_ARGS);
  return validatePolishedOutput(output, mode, hasInternalFragments);
}

const DELTA_PROMPT_INSTRUCTIONS = `You are writing the "changes since the previous prerelease" section of a prerelease notes file for SubMiner, an Electron app for Japanese sentence mining.

You will receive changelog fragment diffs between the previous prerelease tag and the current build. Fragments are engineer-written release-note sources; a fragment diff is a proxy for what changed, not proof of a behavior change.

Rules:

1. Output Markdown bullets ONLY. No headings, no preamble, no commentary. Every line must be a top-level "- " bullet or an indented nested bullet.
2. Describe only what changed for users between the two prerelease builds, in user-facing language. Drop implementation jargon, file paths, and PR numbers.
3. ADDED fragments describe changes that are new in this build; summarize them.
4. MODIFIED fragments include BEFORE and AFTER content. Describe only the behavioral difference between them. If the edit is editorial (rewording, deduplication, reformatting, reconciling stale phrasing) with no user-visible behavior change, omit it entirely.
5. DELETED fragments mean the described change was removed or reverted before this build; say so explicitly.
6. Keep bullets short and concrete. Use nested bullets sparingly.
7. Do not invent changes. Every bullet must be grounded in the diffs.
8. If no bullet survives rules 2-5, output exactly this single line:
   - No user-facing changes since PREVIOUS_TAG.

The input begins below.

`;

function serializeFragmentDeltaForPrompt(
  delta: FragmentDeltaEntry[],
  version: string,
  previousTag: string,
): string {
  const header = [`VERSION: ${version}`, `PREVIOUS_TAG: ${previousTag}`];
  const blocks = delta.map((entry) => {
    if (entry.status === 'added') {
      return [`ADDED FRAGMENT ${entry.path}`, entry.after ?? ''].join('\n');
    }
    if (entry.status === 'deleted') {
      return [`DELETED FRAGMENT ${entry.path}`, entry.before ?? ''].join('\n');
    }
    return [
      `MODIFIED FRAGMENT ${entry.path}`,
      'BEFORE:',
      entry.before ?? '',
      'AFTER:',
      entry.after ?? '',
    ].join('\n');
  });
  return [...header, '', ...blocks].join('\n\n');
}

function validateDeltaOutput(output: string): string {
  const trimmed = output.trim();
  if (!trimmed) {
    throw new Error('claude returned empty output for the prerelease delta section.');
  }
  const invalidLine = trimmed.split(/\r?\n/).find((line) => line.trim() && !/^\s*- /.test(line));
  if (invalidLine !== undefined) {
    throw new Error(
      `claude delta output must contain only Markdown bullets. Offending line:\n${invalidLine}`,
    );
  }
  return trimmed;
}

function buildDeltaSectionWithClaude(
  delta: FragmentDeltaEntry[],
  options: { version: string; previousTag: string; deps?: ChangelogFsDeps },
): string {
  const runClaude = options.deps?.runClaude ?? defaultRunClaude;
  const prompt =
    DELTA_PROMPT_INSTRUCTIONS.replace('PREVIOUS_TAG', options.previousTag) +
    serializeFragmentDeltaForPrompt(delta, options.version, options.previousTag);
  return validateDeltaOutput(runClaude(prompt, CLAUDE_CLI_ARGS));
}

function stripDetailsBlocks(body: string): string {
  return body.replace(/<details>[\s\S]*?<\/details>\s*/gm, '').trim();
}

function buildReleaseSection(
  version: string,
  date: string,
  fragments: ChangeFragment[],
  deps?: ChangelogFsDeps,
): string {
  if (fragments.length === 0) {
    throw new Error('No changelog fragments found in changes/.');
  }

  const polished = polishFragmentsWithClaude(fragments, {
    mode: 'changelog',
    version,
    date,
    deps,
  });
  return [`## v${version} (${date})`, '', polished, ''].join('\n');
}

function ensureChangelogHeader(existingChangelog: string): string {
  const trimmed = existingChangelog.trim();
  if (!trimmed) {
    return `${CHANGELOG_HEADER}\n`;
  }
  if (trimmed.startsWith(CHANGELOG_HEADER)) {
    return `${trimmed}\n`;
  }
  return `${CHANGELOG_HEADER}\n\n${trimmed}\n`;
}

function prependReleaseSection(
  existingChangelog: string,
  releaseSection: string,
  version: string,
): string {
  const normalizedExisting = ensureChangelogHeader(existingChangelog);
  if (extractReleaseSectionBody(normalizedExisting, version) !== null) {
    throw new Error(`CHANGELOG already contains a section for v${version}.`);
  }

  const withoutHeader = normalizedExisting.replace(/^# Changelog\s*/, '').trimStart();
  const body = [releaseSection.trimEnd(), withoutHeader.trimEnd()].filter(Boolean).join('\n\n');
  return `${CHANGELOG_HEADER}\n\n${body}\n`;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function extractReleaseSectionBody(changelog: string, version: string): string | null {
  const headingPattern = new RegExp(
    `^## v${escapeRegExp(normalizeVersion(version))} \\([^\\n]+\\)$`,
    'm',
  );
  const headingMatch = headingPattern.exec(changelog);
  if (!headingMatch) {
    return null;
  }

  const bodyStart = headingMatch.index + headingMatch[0].length + 1;
  const remaining = changelog.slice(bodyStart);
  const nextHeadingMatch = /^## /m.exec(remaining);
  const body = nextHeadingMatch ? remaining.slice(0, nextHeadingMatch.index) : remaining;
  return body.trim();
}

export function resolveChangelogOutputPaths(options?: { cwd?: string }): string[] {
  const cwd = options?.cwd ?? process.cwd();
  return [path.join(cwd, 'CHANGELOG.md')];
}

function renderReleaseNotes(
  changes: string,
  options?: {
    disclaimer?: string;
    contributions?: Contribution[];
    contributorSections?: string[];
    metadata?: string[];
    deltaSection?: string[];
  },
): string {
  const prefix = options?.disclaimer ? [options.disclaimer, ''] : [];
  const metadata = options?.metadata?.length ? [...options.metadata, ''] : [];
  const deltaSection = options?.deltaSection?.length ? [...options.deltaSection, ''] : [];
  const contributorSections =
    options?.contributorSections ?? renderContributorsSections(options?.contributions ?? []);
  return [
    ...prefix,
    ...metadata,
    ...deltaSection,
    '## Highlights',
    changes,
    '',
    ...contributorSections,
    '## Installation',
    '',
    'See the README and docs/installation guide for full setup steps.',
    '',
    '## Assets',
    '',
    '- Linux: `SubMiner.AppImage`',
    '- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`',
    '- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`',
    '- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher',
    '',
    'Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.',
    '',
  ].join('\n');
}

function writeReleaseNotesFile(
  cwd: string,
  changes: string,
  deps?: ChangelogFsDeps,
  options?: {
    disclaimer?: string;
    outputPath?: string;
    contributions?: Contribution[];
    contributorSections?: string[];
    metadata?: string[];
    deltaSection?: string[];
  },
): string {
  const mkdirSync = deps?.mkdirSync ?? fs.mkdirSync;
  const writeFileSync = deps?.writeFileSync ?? fs.writeFileSync;
  const releaseNotesPath = path.join(cwd, options?.outputPath ?? RELEASE_NOTES_PATH);

  mkdirSync(path.dirname(releaseNotesPath), { recursive: true });
  writeFileSync(releaseNotesPath, renderReleaseNotes(changes, options), 'utf8');
  return releaseNotesPath;
}

export function writeChangelogArtifacts(options?: ChangelogOptions): {
  deletedFragmentPaths: string[];
  outputPaths: string[];
  releaseNotesPath: string;
} {
  const cwd = options?.cwd ?? process.cwd();
  const existsSync = options?.deps?.existsSync ?? fs.existsSync;
  const mkdirSync = options?.deps?.mkdirSync ?? fs.mkdirSync;
  const readFileSync = options?.deps?.readFileSync ?? fs.readFileSync;
  const rmSync = options?.deps?.rmSync ?? fs.rmSync;
  const writeFileSync = options?.deps?.writeFileSync ?? fs.writeFileSync;
  const log = options?.deps?.log ?? console.log;
  const version = resolveVersion(options ?? {});
  const date = resolveDate(options?.date);
  const fragments = readChangeFragments(cwd, options?.deps);
  const contributions = resolveContributionsForFragments(fragments, cwd, options?.deps);
  const existingChangelogPath = path.join(cwd, 'CHANGELOG.md');
  const existingChangelog = existsSync(existingChangelogPath)
    ? readFileSync(existingChangelogPath, 'utf8')
    : '';
  const outputPaths = resolveChangelogOutputPaths({ cwd });
  const existingReleaseSection = extractReleaseSectionBody(existingChangelog, version);
  if (existingReleaseSection !== null) {
    log(`Existing section found for v${version}; skipping changelog prepend.`);
    for (const fragment of fragments) {
      rmSync(fragment.path);
      log(`Removed ${fragment.path}`);
    }

    const releaseNotesPath = writeReleaseNotesFile(
      cwd,
      stripDetailsBlocks(existingReleaseSection),
      options?.deps,
      { contributions },
    );
    log(`Generated ${releaseNotesPath}`);

    return {
      deletedFragmentPaths: fragments.map((fragment) => fragment.path),
      outputPaths,
      releaseNotesPath,
    };
  }

  const releaseSection = buildReleaseSection(version, date, fragments, options?.deps);
  const nextChangelog = prependReleaseSection(existingChangelog, releaseSection, version);

  for (const outputPath of outputPaths) {
    mkdirSync(path.dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, nextChangelog, 'utf8');
    log(`Updated ${outputPath}`);
  }

  const releaseNotesBody = polishFragmentsWithClaude(fragments, {
    mode: 'release-notes',
    version,
    date,
    deps: options?.deps,
  });
  const releaseNotesPath = writeReleaseNotesFile(cwd, releaseNotesBody, options?.deps, {
    contributions,
  });
  log(`Generated ${releaseNotesPath}`);

  for (const fragment of fragments) {
    rmSync(fragment.path);
    log(`Removed ${fragment.path}`);
  }

  return {
    deletedFragmentPaths: fragments.map((fragment) => fragment.path),
    outputPaths,
    releaseNotesPath,
  };
}

export function writeStableReleaseArtifacts(options?: ChangelogOptions): {
  deletedFragmentPaths: string[];
  docsChangelogPath: string;
  outputPaths: string[];
  releaseNotesPath: string;
} {
  const changelogResult = writeChangelogArtifacts(options);
  const docsChangelogPath = generateDocsChangelog(options);

  return {
    ...changelogResult,
    docsChangelogPath,
  };
}

export function verifyChangelogFragments(options?: ChangelogOptions): void {
  readChangeFragments(options?.cwd ?? process.cwd(), options?.deps);
}

export function verifyChangelogReadyForRelease(options?: ChangelogOptions): void {
  const cwd = options?.cwd ?? process.cwd();
  const readFileSync = options?.deps?.readFileSync ?? fs.readFileSync;
  verifyRequestedVersionMatchesPackageVersion(options ?? {});
  const version = resolveVersion(options ?? {});
  const pendingFragments = resolveFragmentPaths(cwd, options?.deps);
  if (pendingFragments.length > 0) {
    throw new Error(
      `Pending changelog fragments must be released first: ${pendingFragments.join(', ')}`,
    );
  }

  const changelogPath = path.join(cwd, 'CHANGELOG.md');
  if (!(options?.deps?.existsSync ?? fs.existsSync)(changelogPath)) {
    throw new Error(`Missing ${changelogPath}`);
  }

  const changelog = readFileSync(changelogPath, 'utf8');
  if (extractReleaseSectionBody(changelog, version) === null) {
    throw new Error(`Missing CHANGELOG section for v${version}.`);
  }
}

function isFragmentPath(candidate: string): boolean {
  return /^changes\/.+\.md$/u.test(candidate) && !/\/?README\.md$/iu.test(candidate);
}

function isIgnoredPullRequestPath(candidate: string): boolean {
  return (
    candidate === 'CHANGELOG.md' ||
    candidate === 'release/release-notes.md' ||
    candidate === 'AGENTS.md' ||
    candidate === 'README.md' ||
    candidate.startsWith('changes/') ||
    candidate.startsWith('docs/') ||
    candidate.startsWith('.github/')
  );
}

export function verifyPullRequestChangelog(options: PullRequestChangelogOptions): void {
  const labels = (options.changedLabels ?? []).map((label) => label.trim()).filter(Boolean);
  if (labels.includes(SKIP_CHANGELOG_LABEL)) {
    return;
  }

  const normalizedEntries = options.changedEntries
    .map((entry) => ({
      path: entry.path.trim(),
      status: entry.status.trim().toUpperCase(),
    }))
    .filter((entry) => entry.path);
  if (normalizedEntries.length === 0) {
    return;
  }

  const fragmentEntries = normalizedEntries.filter(
    (entry) => entry.status !== 'D' && isFragmentPath(entry.path),
  );
  const hasFragment = fragmentEntries.length > 0;
  const requiresFragment = normalizedEntries.some((entry) => !isIgnoredPullRequestPath(entry.path));

  if (requiresFragment && !hasFragment) {
    throw new Error(
      `This pull request changes release-relevant files and requires a reconciled changelog fragment under changes/ or the ${SKIP_CHANGELOG_LABEL} label. Before adding a new fragment, update the existing PR fragment when the new work modifies, fixes, or supersedes behavior already described there.`,
    );
  }
}

function resolveChangedPathsFromGit(
  cwd: string,
  baseRef: string,
  headRef: string,
): Array<{ path: string; status: string }> {
  const output = execFileSync('git', ['diff', '--name-status', `${baseRef}...${headRef}`], {
    cwd,
    encoding: 'utf8',
  });
  return output
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [status = '', ...paths] = line.split(/\s+/);
      return {
        path: paths[paths.length - 1] ?? '',
        status,
      };
    })
    .filter((entry) => entry.path);
}

const DOCS_CHANGELOG_PATH = path.join('docs-site', 'changelog.md');

type VersionSection = {
  version: string;
  date: string;
  minor: string;
  body: string;
};

function parseVersionSections(changelog: string): VersionSection[] {
  const sectionPattern = /^## v(\d+\.\d+\.\d+) \((\d{4}-\d{2}-\d{2})\)$/gm;
  const sections: VersionSection[] = [];
  let match: RegExpExecArray | null;

  while ((match = sectionPattern.exec(changelog)) !== null) {
    const version = match[1]!;
    const date = match[2]!;
    const minor = version.replace(/\.\d+$/, '');
    const headingEnd = match.index + match[0].length;
    sections.push({ version, date, minor, body: '' });

    if (sections.length > 1) {
      const prev = sections[sections.length - 2]!;
      prev.body = changelog.slice(prev.body as unknown as number, match.index).trim();
    }
    (sections[sections.length - 1] as { body: unknown }).body = headingEnd;
  }

  if (sections.length > 0) {
    const last = sections[sections.length - 1]!;
    last.body = changelog.slice(last.body as unknown as number).trim();
  }

  return sections;
}

export function generateDocsChangelog(options?: Pick<ChangelogOptions, 'cwd' | 'deps'>): string {
  const cwd = options?.cwd ?? process.cwd();
  const readFileSync = options?.deps?.readFileSync ?? fs.readFileSync;
  const writeFileSync = options?.deps?.writeFileSync ?? fs.writeFileSync;
  const log = options?.deps?.log ?? console.log;

  const changelogPath = path.join(cwd, 'CHANGELOG.md');
  const changelog = readFileSync(changelogPath, 'utf8');
  const sections = parseVersionSections(changelog);

  if (sections.length === 0) {
    throw new Error('No version sections found in CHANGELOG.md');
  }

  const currentMinor = sections[0]!.minor;
  const currentSections = sections.filter((s) => s.minor === currentMinor);
  const olderSections = sections.filter((s) => s.minor !== currentMinor);

  const lines: string[] = ['# Changelog', ''];

  for (const section of currentSections) {
    const body = section.body.replace(/^### (.+)$/gm, '**$1**');
    lines.push(`## v${section.version} (${section.date})`, '', body, '');
  }

  if (olderSections.length > 0) {
    lines.push('## Previous Versions', '');

    const minorGroups = new Map<string, VersionSection[]>();
    for (const section of olderSections) {
      const group = minorGroups.get(section.minor) ?? [];
      group.push(section);
      minorGroups.set(section.minor, group);
    }

    for (const [minor, group] of minorGroups) {
      lines.push('<details>', `<summary>v${minor}.x</summary>`, '');
      for (const section of group) {
        const htmlBody = section.body.replace(/^### (.+)$/gm, '**$1**');
        lines.push(`<h2>v${section.version} (${section.date})</h2>`, '', htmlBody, '');
      }
      lines.push('</details>', '');
    }
  }

  const output =
    lines
      .join('\n')
      .replace(/\n{3,}/g, '\n\n')
      .trimEnd() + '\n';
  const outputPath = path.join(cwd, DOCS_CHANGELOG_PATH);
  writeFileSync(outputPath, output, 'utf8');
  log(`Generated ${outputPath}`);

  return outputPath;
}

export function writeReleaseNotesForVersion(options?: ChangelogOptions): string {
  const cwd = options?.cwd ?? process.cwd();
  const existsSync = options?.deps?.existsSync ?? fs.existsSync;
  const readFileSync = options?.deps?.readFileSync ?? fs.readFileSync;
  const version = resolveVersion(options ?? {});
  const changelogPath = path.join(cwd, 'CHANGELOG.md');
  const changelog = readFileSync(changelogPath, 'utf8');
  const changes = extractReleaseSectionBody(changelog, version);

  if (changes === null) {
    throw new Error(`Missing CHANGELOG section for v${version}.`);
  }

  const releaseNotesPath = path.join(cwd, RELEASE_NOTES_PATH);
  const contributorSections = existsSync(releaseNotesPath)
    ? extractContributorSections(readFileSync(releaseNotesPath, 'utf8'))
    : [];

  return writeReleaseNotesFile(cwd, stripDetailsBlocks(changes), options?.deps, {
    contributorSections,
  });
}

export function writePrereleaseNotesForVersion(options?: ChangelogOptions): string {
  verifyRequestedVersionMatchesPackageVersion(options ?? {});

  const cwd = options?.cwd ?? process.cwd();
  const existsSync = options?.deps?.existsSync ?? fs.existsSync;
  const readFileSync = options?.deps?.readFileSync ?? fs.readFileSync;
  const version = resolveVersion(options ?? {});
  if (!isSupportedPrereleaseVersion(version)) {
    throw new Error(
      `Unsupported prerelease version (${version}). Expected x.y.z-beta.N or x.y.z-rc.N.`,
    );
  }

  const fragments = readChangeFragments(cwd, options?.deps);
  if (fragments.length === 0) {
    throw new Error('No changelog fragments found in changes/.');
  }

  const listPrereleaseTags = options?.deps?.listPrereleaseTags ?? defaultListPrereleaseTags;
  const previousTag = selectPreviousPrereleaseTag(
    listPrereleaseTags(cwd, resolvePrereleaseBaseVersion(version)),
    version,
  );

  // Later betas/RCs get a "Changes since <previous tag>" section on top of the
  // cumulative Highlights, generated from the fragment diff between the
  // previous prerelease tag and the working tree.
  let deltaSection: string[] = [];
  if (previousTag) {
    const resolveFragmentDelta = options?.deps?.resolveFragmentDelta ?? defaultResolveFragmentDelta;
    const delta = resolveFragmentDelta(cwd, previousTag);
    const deltaBody =
      delta.length === 0
        ? `- No changelog fragment changes since ${previousTag}; this build contains packaging or internal-only updates.`
        : buildDeltaSectionWithClaude(delta, { version, previousTag, deps: options?.deps });
    deltaSection = [`${DELTA_SECTION_HEADING_PREFIX}${previousTag}`, '', deltaBody];
  }

  const prereleaseNotesPath = path.join(cwd, PRERELEASE_NOTES_PATH);
  const existingReleaseNotes = existsSync(prereleaseNotesPath)
    ? resolveReusablePrereleaseNotes(readFileSync(prereleaseNotesPath, 'utf8'), version)
    : undefined;
  const changes = polishFragmentsWithClaude(fragments, {
    mode: 'release-notes',
    version,
    existingReleaseNotes,
    deps: options?.deps,
  });
  const contributions = resolveContributionsForFragments(fragments, cwd, options?.deps);
  return writeReleaseNotesFile(cwd, changes, options?.deps, {
    disclaimer:
      '> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.',
    outputPath: PRERELEASE_NOTES_PATH,
    contributions,
    metadata: [renderPrereleaseVersionMarker(version, previousTag)],
    deltaSection,
  });
}

// CI gate: the committed prerelease notes must carry a marker generated for
// exactly the version being tagged, so stale beta.N-1 notes can't ship.
export function verifyPrereleaseNotesMatchVersion(options?: ChangelogOptions): void {
  verifyRequestedVersionMatchesPackageVersion(options ?? {});

  const cwd = options?.cwd ?? process.cwd();
  const existsSync = options?.deps?.existsSync ?? fs.existsSync;
  const readFileSync = options?.deps?.readFileSync ?? fs.readFileSync;
  const version = resolveVersion(options ?? {});
  if (!isSupportedPrereleaseVersion(version)) {
    throw new Error(
      `Unsupported prerelease version (${version}). Expected x.y.z-beta.N or x.y.z-rc.N.`,
    );
  }

  const prereleaseNotesPath = path.join(cwd, PRERELEASE_NOTES_PATH);
  if (!existsSync(prereleaseNotesPath)) {
    throw new Error(
      `Missing ${prereleaseNotesPath}. Run 'bun run changelog:prerelease-notes --version ${version}' and commit the file before tagging.`,
    );
  }

  const markerVersion = extractPrereleaseVersionMarker(readFileSync(prereleaseNotesPath, 'utf8'));
  if (markerVersion !== version) {
    throw new Error(
      `release/prerelease-notes.md was generated for ${markerVersion ?? 'an unknown version (missing or legacy prerelease-version marker)'} but this release is ${version}. Rerun 'bun run changelog:prerelease-notes --version ${version}' and commit the result.`,
    );
  }
}

function parseCliArgs(argv: string[]): {
  baseRef?: string;
  cwd?: string;
  date?: string;
  headRef?: string;
  labels?: string;
  version?: string;
} {
  const parsed: {
    baseRef?: string;
    cwd?: string;
    date?: string;
    headRef?: string;
    labels?: string;
    version?: string;
  } = {};

  for (let index = 0; index < argv.length; index += 1) {
    const current = argv[index];
    const next = argv[index + 1];

    if (current === '--cwd' && next) {
      parsed.cwd = next;
      index += 1;
      continue;
    }

    if (current === '--date' && next) {
      parsed.date = next;
      index += 1;
      continue;
    }

    if (current === '--version' && next) {
      parsed.version = next;
      index += 1;
      continue;
    }

    if (current === '--base-ref' && next) {
      parsed.baseRef = next;
      index += 1;
      continue;
    }

    if (current === '--head-ref' && next) {
      parsed.headRef = next;
      index += 1;
      continue;
    }

    if (current === '--labels' && next) {
      parsed.labels = next;
      index += 1;
    }
  }

  return parsed;
}

function main(): void {
  const [command = 'build', ...argv] = process.argv.slice(2);
  const options = parseCliArgs(argv);

  if (command === 'build') {
    writeChangelogArtifacts(options);
    return;
  }

  if (command === 'build-release') {
    writeStableReleaseArtifacts(options);
    return;
  }

  if (command === 'check') {
    verifyChangelogReadyForRelease(options);
    return;
  }

  if (command === 'lint') {
    verifyChangelogFragments(options);
    return;
  }

  if (command === 'pr-check') {
    verifyChangelogFragments(options);
    verifyPullRequestChangelog({
      changedLabels: options.labels?.split(',') ?? [],
      changedEntries: resolveChangedPathsFromGit(
        options.cwd ?? process.cwd(),
        options.baseRef ?? 'origin/main',
        options.headRef ?? 'HEAD',
      ),
    });
    return;
  }

  if (command === 'release-notes') {
    writeReleaseNotesForVersion(options);
    return;
  }

  if (command === 'prerelease-notes') {
    writePrereleaseNotesForVersion(options);
    return;
  }

  if (command === 'check-prerelease-notes') {
    verifyPrereleaseNotesMatchVersion(options);
    return;
  }

  if (command === 'docs') {
    generateDocsChangelog(options);
    return;
  }

  throw new Error(`Unknown changelog command: ${command}`);
}

if (require.main === module) {
  main();
}
