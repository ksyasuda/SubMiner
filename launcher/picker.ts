import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { spawnSync } from "node:child_process";
import type { LogLevel, JellyfinSessionConfig, JellyfinLibraryEntry, JellyfinItemEntry, JellyfinGroupEntry } from "./types.js";
import { VIDEO_EXTENSIONS, ROFI_THEME_FILE } from "./types.js";
import { log, fail } from "./log.js";
import { commandExists, realpathMaybe } from "./util.js";

export function escapeShellSingle(value: string): string {
  return `'${value.replace(/'/g, `'\\''`)}'`;
}

export function showRofiFlatMenu(
  items: string[],
  prompt: string,
  initialQuery = "",
  themePath: string | null = null,
): string {
  const args = [
    "-dmenu",
    "-i",
    "-matching",
    "fuzzy",
    "-p",
    prompt,
  ];
  if (themePath) {
    args.push("-theme", themePath);
  } else {
    args.push(
      "-theme-str",
      'configuration { font: "Noto Sans CJK JP Regular 8";}',
    );
  }
  if (initialQuery.trim().length > 0) {
    args.push("-filter", initialQuery.trim());
  }
  const result = spawnSync(
    "rofi",
    args,
    {
      input: `${items.join("\n")}\n`,
      encoding: "utf8",
      stdio: ["pipe", "pipe", "ignore"],
    },
  );
  if (result.error) {
    fail(formatPickerLaunchError("rofi", result.error as NodeJS.ErrnoException));
  }
  return (result.stdout || "").trim();
}

export function showFzfFlatMenu(
  lines: string[],
  prompt: string,
  previewCommand: string,
  initialQuery = "",
): string {
  const args = [
    "--ansi",
    "--reverse",
    "--ignore-case",
    `--prompt=${prompt}`,
    "--delimiter=\t",
    "--with-nth=2",
    "--preview-window=right:50%:wrap",
    "--preview",
    previewCommand,
  ];
  if (initialQuery.trim().length > 0) {
    args.push("--query", initialQuery.trim());
  }
  const result = spawnSync(
    "fzf",
    args,
    {
      input: `${lines.join("\n")}\n`,
      encoding: "utf8",
      stdio: ["pipe", "pipe", "inherit"],
    },
  );
  if (result.error) {
    fail(formatPickerLaunchError("fzf", result.error as NodeJS.ErrnoException));
  }
  return (result.stdout || "").trim();
}

export function parseSelectionId(selection: string): string {
  if (!selection) return "";
  const tab = selection.indexOf("\t");
  if (tab === -1) return "";
  return selection.slice(0, tab);
}

export function parseSelectionLabel(selection: string): string {
  const tab = selection.indexOf("\t");
  if (tab === -1) return selection;
  return selection.slice(tab + 1);
}

function fuzzySubsequenceMatch(haystack: string, needle: string): boolean {
  if (!needle) return true;
  let j = 0;
  for (let i = 0; i < haystack.length && j < needle.length; i += 1) {
    if (haystack[i] === needle[j]) j += 1;
  }
  return j === needle.length;
}

function matchesMenuQuery(label: string, query: string): boolean {
  const normalizedQuery = query.trim().toLowerCase();
  if (!normalizedQuery) return true;
  const target = label.toLowerCase();
  const tokens = normalizedQuery.split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return true;
  return tokens.every((token) => fuzzySubsequenceMatch(target, token));
}

export async function promptOptionalJellyfinSearch(
  useRofi: boolean,
  themePath: string | null = null,
): Promise<string> {
  if (useRofi && commandExists("rofi")) {
    const rofiArgs = [
      "-dmenu",
      "-i",
      "-p",
      "Jellyfin Search (optional)",
    ];
    if (themePath) {
      rofiArgs.push("-theme", themePath);
    } else {
      rofiArgs.push(
        "-theme-str",
        'configuration { font: "Noto Sans CJK JP Regular 8";}',
      );
    }
    const result = spawnSync(
      "rofi",
      rofiArgs,
      {
        input: "\n",
        encoding: "utf8",
        stdio: ["pipe", "pipe", "ignore"],
      },
    );
    if (result.error) return "";
    return (result.stdout || "").trim();
  }

  if (!process.stdin.isTTY || !process.stdout.isTTY) return "";

  process.stdout.write("Jellyfin search term (optional, press Enter to skip): ");
  const chunks: Buffer[] = [];
  return await new Promise<string>((resolve) => {
    const onData = (data: Buffer) => {
      const line = data.toString("utf8");
      if (line.includes("\n") || line.includes("\r")) {
        chunks.push(Buffer.from(line, "utf8"));
        process.stdin.off("data", onData);
        const text = Buffer.concat(chunks).toString("utf8").trim();
        resolve(text);
        return;
      }
      chunks.push(data);
    };
    process.stdin.on("data", onData);
  });
}

interface RofiIconEntry {
  label: string;
  iconPath?: string;
}

function showRofiIconMenu(
  entries: RofiIconEntry[],
  prompt: string,
  initialQuery = "",
  themePath: string | null = null,
): number {
  if (entries.length === 0) return -1;
  const rofiArgs = ["-dmenu", "-i", "-show-icons", "-format", "i", "-p", prompt];
  if (initialQuery) rofiArgs.push("-filter", initialQuery);
  if (themePath) {
    rofiArgs.push("-theme", themePath);
  } else {
    rofiArgs.push(
      "-theme-str",
      'configuration { font: "Noto Sans CJK JP Regular 8";}',
    );
  }

  const lines = entries.map((entry) =>
    entry.iconPath
      ? `${entry.label}\u0000icon\u001f${entry.iconPath}`
      : entry.label
  );
  const result = spawnSync(
    "rofi",
    rofiArgs,
    {
      input: `${lines.join("\n")}\n`,
      encoding: "utf8",
      stdio: ["pipe", "pipe", "ignore"],
    },
  );
  if (result.error) return -1;
  const out = (result.stdout || "").trim();
  if (!out) return -1;
  const idx = Number.parseInt(out, 10);
  return Number.isFinite(idx) ? idx : -1;
}

export function pickLibrary(
  session: JellyfinSessionConfig,
  libraries: JellyfinLibraryEntry[],
  useRofi: boolean,
  ensureIcon: (session: JellyfinSessionConfig, id: string) => string | null,
  initialQuery = "",
  themePath: string | null = null,
): string {
  const visibleLibraries = initialQuery.trim().length > 0
    ? libraries.filter((lib) =>
      matchesMenuQuery(`${lib.name} ${lib.kind}`, initialQuery)
    )
    : libraries;
  if (visibleLibraries.length === 0) fail("No Jellyfin libraries found.");

  if (useRofi) {
    const entries = visibleLibraries.map((lib) => ({
      label: `${lib.name} [${lib.kind}]`,
      iconPath: ensureIcon(session, lib.id) || undefined,
    }));
    const idx = showRofiIconMenu(
      entries,
      "Jellyfin Library",
      initialQuery,
      themePath,
    );
    return idx >= 0 ? visibleLibraries[idx].id : "";
  }

  const lines = visibleLibraries.map(
    (lib) => `${lib.id}\t${lib.name} [${lib.kind}]`,
  );
  const preview = commandExists("chafa") && commandExists("curl")
    ? `
id={1}
url=${escapeShellSingle(session.serverUrl)}/Items/$id/Images/Primary?maxHeight=720\\&quality=85\\&api_key=${escapeShellSingle(session.accessToken)}
curl -fsSL "$url" 2>/dev/null | chafa --format=symbols --symbols=vhalf+wide --size=${"${FZF_PREVIEW_COLUMNS}"}x${"${FZF_PREVIEW_LINES}"} - 2>/dev/null
`.trim()
    : 'echo "Install curl + chafa for image preview"';

  const picked = showFzfFlatMenu(
    lines,
    "Jellyfin Library: ",
    preview,
    initialQuery,
  );
  return parseSelectionId(picked);
}

export function pickItem(
  session: JellyfinSessionConfig,
  items: JellyfinItemEntry[],
  useRofi: boolean,
  ensureIcon: (session: JellyfinSessionConfig, id: string) => string | null,
  initialQuery = "",
  themePath: string | null = null,
): string {
  const visibleItems = initialQuery.trim().length > 0
    ? items.filter((item) => matchesMenuQuery(item.display, initialQuery))
    : items;
  if (visibleItems.length === 0) fail("No playable Jellyfin items found.");

  if (useRofi) {
    const entries = visibleItems.map((item) => ({
      label: item.display,
      iconPath: ensureIcon(session, item.id) || undefined,
    }));
    const idx = showRofiIconMenu(
      entries,
      "Jellyfin Item",
      initialQuery,
      themePath,
    );
    return idx >= 0 ? visibleItems[idx].id : "";
  }

  const lines = visibleItems.map((item) => `${item.id}\t${item.display}`);
  const preview = commandExists("chafa") && commandExists("curl")
    ? `
id={1}
url=${escapeShellSingle(session.serverUrl)}/Items/$id/Images/Primary?maxHeight=720\\&quality=85\\&api_key=${escapeShellSingle(session.accessToken)}
curl -fsSL "$url" 2>/dev/null | chafa --format=symbols --symbols=vhalf+wide --size=${"${FZF_PREVIEW_COLUMNS}"}x${"${FZF_PREVIEW_LINES}"} - 2>/dev/null
`.trim()
    : 'echo "Install curl + chafa for image preview"';

  const picked = showFzfFlatMenu(lines, "Jellyfin Item: ", preview, initialQuery);
  return parseSelectionId(picked);
}

export function pickGroup(
  session: JellyfinSessionConfig,
  groups: JellyfinGroupEntry[],
  useRofi: boolean,
  ensureIcon: (session: JellyfinSessionConfig, id: string) => string | null,
  initialQuery = "",
  themePath: string | null = null,
): string {
  const visibleGroups = initialQuery.trim().length > 0
    ? groups.filter((group) => matchesMenuQuery(group.display, initialQuery))
    : groups;
  if (visibleGroups.length === 0) return "";

  if (useRofi) {
    const entries = visibleGroups.map((group) => ({
      label: group.display,
      iconPath: ensureIcon(session, group.id) || undefined,
    }));
    const idx = showRofiIconMenu(
      entries,
      "Jellyfin Anime/Folder",
      initialQuery,
      themePath,
    );
    return idx >= 0 ? visibleGroups[idx].id : "";
  }

  const lines = visibleGroups.map((group) => `${group.id}\t${group.display}`);
  const preview = commandExists("chafa") && commandExists("curl")
    ? `
id={1}
url=${escapeShellSingle(session.serverUrl)}/Items/$id/Images/Primary?maxHeight=720\\&quality=85\\&api_key=${escapeShellSingle(session.accessToken)}
curl -fsSL "$url" 2>/dev/null | chafa --format=symbols --symbols=vhalf+wide --size=${"${FZF_PREVIEW_COLUMNS}"}x${"${FZF_PREVIEW_LINES}"} - 2>/dev/null
`.trim()
    : 'echo "Install curl + chafa for image preview"';

  const picked = showFzfFlatMenu(
    lines,
    "Jellyfin Anime/Folder: ",
    preview,
    initialQuery,
  );
  return parseSelectionId(picked);
}

export function formatPickerLaunchError(
  picker: "rofi" | "fzf",
  error: NodeJS.ErrnoException,
): string {
  if (error.code === "ENOENT") {
    return picker === "rofi"
      ? "rofi not found. Install rofi or use --no-rofi to use fzf."
      : "fzf not found. Install fzf or use --rofi to use rofi.";
  }
  return `Failed to launch ${picker}: ${error.message}`;
}

export function collectVideos(dir: string, recursive: boolean): string[] {
  const root = path.resolve(dir);
  const out: string[] = [];

  const walk = (current: string): void => {
    let entries: fs.Dirent[];
    try {
      entries = fs.readdirSync(current, { withFileTypes: true });
    } catch {
      return;
    }

    for (const entry of entries) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) {
        if (recursive) walk(full);
        continue;
      }
      if (!entry.isFile()) continue;
      const ext = path.extname(entry.name).slice(1).toLowerCase();
      if (VIDEO_EXTENSIONS.has(ext)) out.push(full);
    }
  };

  walk(root);
  return out.sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }),
  );
}

export function buildRofiMenu(
  videos: string[],
  dir: string,
  recursive: boolean,
): Buffer {
  const chunks: Buffer[] = [];
  for (const video of videos) {
    const display = recursive
      ? path.relative(dir, video)
      : path.basename(video);
    const line = `${display}\0icon\x1fthumbnail://${video}\n`;
    chunks.push(Buffer.from(line, "utf8"));
  }
  return Buffer.concat(chunks);
}

export function findRofiTheme(scriptPath: string): string | null {
  const envTheme = process.env.SUBMINER_ROFI_THEME;
  if (envTheme && fs.existsSync(envTheme)) return envTheme;

  const scriptDir = path.dirname(realpathMaybe(scriptPath));
  const candidates: string[] = [];

  if (process.platform === "darwin") {
    candidates.push(
      path.join(
        os.homedir(),
        "Library/Application Support/SubMiner/themes",
        ROFI_THEME_FILE,
      ),
    );
  } else {
    const xdgDataHome =
      process.env.XDG_DATA_HOME || path.join(os.homedir(), ".local/share");
    candidates.push(path.join(xdgDataHome, "SubMiner/themes", ROFI_THEME_FILE));
    candidates.push(
      path.join("/usr/local/share/SubMiner/themes", ROFI_THEME_FILE),
    );
    candidates.push(path.join("/usr/share/SubMiner/themes", ROFI_THEME_FILE));
  }

  candidates.push(path.join(scriptDir, ROFI_THEME_FILE));

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) return candidate;
  }

  return null;
}

export function showRofiMenu(
  videos: string[],
  dir: string,
  recursive: boolean,
  scriptPath: string,
  logLevel: LogLevel,
): string {
  const args = [
    "-dmenu",
    "-i",
    "-p",
    "Select Video ",
    "-show-icons",
    "-theme-str",
    'configuration { font: "Noto Sans CJK JP Regular 8";}',
  ];

  const theme = findRofiTheme(scriptPath);
  if (theme) {
    args.push("-theme", theme);
  } else {
    log(
      "warn",
      logLevel,
      "Rofi theme not found; using rofi defaults (set SUBMINER_ROFI_THEME to override)",
    );
  }

  const result = spawnSync("rofi", args, {
    input: buildRofiMenu(videos, dir, recursive),
    encoding: "utf8",
    stdio: ["pipe", "pipe", "ignore"],
  });
  if (result.error) {
    fail(
      formatPickerLaunchError("rofi", result.error as NodeJS.ErrnoException),
    );
  }

  const selection = (result.stdout || "").trim();
  if (!selection) return "";
  return path.join(dir, selection);
}

export function buildFzfMenu(videos: string[]): string {
  return videos.map((video) => `${path.basename(video)}\t${video}`).join("\n");
}

export function showFzfMenu(videos: string[]): string {
  const chafaFormat = process.env.TMUX
    ? "--format=symbols --symbols=vhalf+wide --color-space=din99d"
    : "--format=kitty";

  const previewCmd = commandExists("chafa")
    ? `
video={2}
thumb_dir="$HOME/.cache/thumbnails/large"
video_uri="file://$(realpath "$video")"
if command -v md5sum >/dev/null 2>&1; then
  thumb_hash=$(echo -n "$video_uri" | md5sum | cut -d' ' -f1)
else
  thumb_hash=$(echo -n "$video_uri" | md5 -q)
fi
thumb_path="$thumb_dir/$thumb_hash.png"

get_thumb() {
  if [[ -f "$thumb_path" ]]; then
    echo "$thumb_path"
  elif command -v ffmpegthumbnailer >/dev/null 2>&1; then
    tmp="/tmp/subminer-preview.jpg"
    ffmpegthumbnailer -i "$video" -o "$tmp" -s 512 -q 5 2>/dev/null && echo "$tmp"
  elif command -v ffmpeg >/dev/null 2>&1; then
    tmp="/tmp/subminer-preview.jpg"
    ffmpeg -y -i "$video" -ss 00:00:05 -vframes 1 -vf "scale=512:-1" "$tmp" 2>/dev/null && echo "$tmp"
  fi
}

thumb=$(get_thumb)
[[ -n "$thumb" ]] && chafa ${chafaFormat} --size=${"${FZF_PREVIEW_COLUMNS}"}x${"${FZF_PREVIEW_LINES}"} "$thumb" 2>/dev/null
`.trim()
    : 'echo "Install chafa for thumbnail preview"';

  const result = spawnSync(
    "fzf",
    [
      "--ansi",
      "--reverse",
      "--prompt=Select Video: ",
      "--delimiter=\t",
      "--with-nth=1",
      "--preview-window=right:50%:wrap",
      "--preview",
      previewCmd,
    ],
    {
      input: buildFzfMenu(videos),
      encoding: "utf8",
      stdio: ["pipe", "pipe", "inherit"],
    },
  );
  if (result.error) {
    fail(formatPickerLaunchError("fzf", result.error as NodeJS.ErrnoException));
  }

  const selection = (result.stdout || "").trim();
  if (!selection) return "";
  const tabIndex = selection.indexOf("\t");
  if (tabIndex === -1) return "";
  return selection.slice(tabIndex + 1);
}
