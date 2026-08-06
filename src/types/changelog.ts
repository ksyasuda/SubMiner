export type ChangelogItem = {
  /** Bullet text with inline markdown preserved. */
  text: string;
  /** Nested bullets, as written with indentation in CHANGELOG.md. */
  children: ChangelogItem[];
};

export type ChangelogSection = {
  /** Section heading as written in CHANGELOG.md, e.g. "Added", "Fixed". */
  heading: string;
  items: ChangelogItem[];
  /** True when the section lived inside the collapsed "Internal changes" block. */
  internal: boolean;
};

export type ChangelogEntry = {
  version: string;
  date: string;
  /** major.minor of `version`, used to decide which entries render expanded. */
  groupKey: string;
  sections: ChangelogSection[];
};

export type ChangelogSourceKind = 'remote' | 'bundled';

export type ChangelogSnapshot = {
  entries: ChangelogEntry[];
  /** Version this build reports, e.g. "0.19.2". */
  installedVersion: string;
  /** Newest version present in the fetched changelog, or null when empty. */
  latestVersion: string | null;
  /** Group key whose entries should start expanded. */
  expandedGroupKey: string | null;
  source: ChangelogSourceKind;
  /** Release tag the remote changelog was read from, when known. */
  releaseTag?: string;
  /** Populated when the remote fetch failed and a fallback was used. */
  warning?: string;
  /** Populated when no changelog could be loaded at all. */
  error?: string;
};
