# Anime Browser

Search anime sources, pick an episode, and play it in mpv with SubMiner's overlay
and mining tools attached — the same way a local file or a Jellyfin stream works.

Open it with `subminer anime`, with `SubMiner.AppImage --anime`, or from
**Browse Anime** in the tray menu. The window stays open while you watch, so you
can queue the next episode without reopening it.

During playback, `Ctrl+Alt+A` toggles the same browser as a modal inside the mpv
player bounds. It uses a dedicated modal surface, so it stays above fullscreen
playback and closes like the other in-player tools. Toggling it off keeps its
current page and scroll position ready for the next toggle. The standalone
window and the modal keep their own search, selected source, tab, and scroll
state, so using one does not replace or cancel what you were doing in the other.

Both surfaces use the same playback queue, source configuration, and stats
history. Queue changes and the currently playing episode appear in both
immediately, and watched marks come from the same history that playback and the
stats window update. Closing and reopening the modal therefore picks up progress
made from either browser surface.

While the window is open, SubMiner shows a tray icon and — on macOS — appears
in the Cmd+Tab switcher and the Dock (macOS ties the two together), so you can
flip between the browser and mpv. SubMiner normally hides itself from the Dock
because the subtitle overlay needs that to float above fullscreen video; it
hides again when the window closes during playback.

Launching an episode starts a full SubMiner playback session, the same as
playing a local file: the overlay and mining tools attach, and the tray icon
stays available. In standalone `subminer anime` mode, closing the window while
a video is playing leaves playback running — reopen the browser from the tray
(**Browse Anime**). The app only exits with the window when nothing is playing.

## How it works

SubMiner does not implement any anime source itself. It runs **Aniyomi extension
APKs** through a bundled JVM sidecar ([M-Extension-Server][mes]), asks the
selected extension to resolve an episode, and hands the resulting URL to mpv.

```
extension APK  →  bridge (JVM)  →  { url, headers }  →  mpv  →  SubMiner overlay
```

Because the extension resolves the stream, whichever sources you install decide
what is available. SubMiner only hosts them.

## Installing extensions

**SubMiner ships no extension repositories and bundles no sources.** There is no
default repository, no suggested list, and no discovery. Until you add one, the
browser has nothing to search — that is deliberate, and it is what keeps SubMiner
a neutral host rather than a distributor.

There are two ways to add extensions.

### From a repository

The window has three tabs — **Browse**, **Extensions**, and **Source settings** —
and each one fills the window, so a long extension list is not squeezed in above
the search results.

Open the **Extensions** tab, paste a repository index URL, and choose
**Add repository**. The URL must be `https` and point at a `.json` index file —
`index.min.json` is the common Aniyomi name, but repositories are free to publish
under another one (for example `video.min.json`). Anything else is rejected
immediately rather than failing later. Everything before the file name is treated
as the repository root, so `.apk` and icon URLs are resolved relative to it.

Extensions your repositories offer but you do not have appear under
**Available**, each with **Install**. Repositories are stored in config under
`anime.repos`, so you can also manage them there and keep them in a dotfile.

Every row carries the extension's icon, as the repository publishes it, so a
site is recognisable before you read the name. A repository row shows its host's
favicon instead. Icons are the only part of a row that is fetched from the
network, and a row whose icon is missing falls back to the first letter of its
name rather than an empty box.

A repository index lists every language it knows about, which is far more than
any one person reads, so the **Available** list has a language chip row above
it. Pick one or more languages to narrow it, or **All** to clear the filter;
picking a language replaces **All** rather than sitting beside it. Rows name the
language in full ("Japanese" rather than `ja`), extensions whose sources span
languages are grouped under **Multi-language**, and the Available heading counts
how many of the offered extensions the filter leaves.

### Managing what is installed

The Extensions tab opens with an **Installed** section listing everything in the
extensions directory, with the sources each one provides and a **Remove**
button. It is built from the directory rather than from a repository, so an
extension you dropped in by hand — or one whose repository you have since
removed — is still listed and still removable. An installed extension borrows
its icon from the catalogue, so one no repository carries shows its monogram.

SubMiner compares the APK's Android version code with the newest build in the
configured repositories. **Update** is clickable only when the repository has a
newer build. A current extension says **Up to date**. If an unusual APK has no
readable version code, its disabled button says **Version unknown**. **Update
all** installs every newer build in one pass and disables itself when there is
nothing to update.

### From a file

Drop Aniyomi `.apk` files into the extensions directory, shown at the top of the
Extensions tab. It defaults to `<userData>/anime-extensions` — on macOS,
`~/Library/Application Support/SubMiner/anime-extensions` — and can be moved with
`anime.extensionsDir`.

A single APK may provide several sources; each appears separately in the
**Source** picker. Extensions that fail to load are listed in the Installed
section with the reason, so a broken APK is visible rather than silently
missing.

An extension that fails to load is skipped rather than blocking the others, so
one bad APK will not hide the rest.

## Searching every source at once

With more than one source installed, the **Source** picker gains an
**All sources** entry. Searching with it selected runs the query against every
installed source at once, and each source's results appear the moment that
source answers — a fast source is on screen while a slow one is still
resolving. The status bar counts sources as they finish
(`Searching… 3/5 sources · 42 results`).

Each cover is labelled with the source it came from, and opening one always
queries that source, whatever the picker says afterwards.

A source that errors is named in the status bar and the rest still show their
results; one extension that needs a login cannot blank the grid. If every
source fails, the first error is shown in full.

Typing a new search while one is still running simply starts over: results
from the superseded search are discarded, even if its sources answer late.

When a source reports another page, **Load more** appears below the covers.
It appends the next page without duplicating entries that already arrived in
the live result stream. A failed next-page request remains available to retry.

Source settings belong to a single extension, so the **Source settings** tab
asks you to pick one while **All sources** is selected.

## Finding an episode, and what you have watched

An episode list can run to hundreds of entries, so the episode header carries a
filter box:

- A number, `12`, keeps that episode. Sources that report no numbers at all
  are still searched by name, so `12` also matches `Episode 12` in a title.
- A range, `12-18`, keeps the episodes between the two, in either order
  (`18-12` reads the same). Episodes the source gave no number are left out of
  a range.
- Anything else is a case-insensitive substring of the episode name, so `beach`
  finds `OVA: Beach Special`.

The counter next to **Episodes** reads `6 of 25` while a filter is applied.
Pressing Escape in the filter box clears it; pressing it anywhere else goes back
to the results grid.

Episodes you have already watched are dimmed and marked `✓ watched`, with a
count in the header. This is not a separate list the browser keeps: it reads the
same stats history the rest of SubMiner writes to, where an episode is marked
watched once a session runs past the completion threshold. Streams are recorded
under a stable per-episode identity, so the mark survives the stream URL
changing between playbacks, and it is the same mark the stats window and
`--mark-watched` use.

Because playback marks an episode partway through the session, the marks
refresh when the browser window comes back to the front: finish an episode in
mpv, switch back, and it is marked. With
[immersion tracking](configuration.md) disabled there is no history to read, so
no episode is marked.

### Marking by hand, and catching up

Right-click an episode for:

- **Mark watched** / **Mark unwatched**: the single episode, whichever way it
  is not already.
- **Mark this and N below watched** / **... unwatched**: that episode and every
  episode listed below it. Sources list newest first, so "below" is the back
  catalogue: right-click the last episode you saw and mark everything down to
  the start, which is how you catch up a series you watched somewhere else.

A filter narrows what you are looking at, not what you mark: a span always
covers the full episode list, and the status bar says how many episodes it
touched. The oldest episode has nothing below it, so it only offers the single
entry. Escape closes the menu.

Marking an episode you have never played creates its stats row so the mark has
somewhere to live. That row carries the same series, season and episode fields
playback would have recorded, and both stats library views join the lifetime
tables, so a manually marked episode does not appear there as watch time you
never spent. Clearing a mark never creates anything.

Marks are written to the stats history, so with immersion tracking disabled
there is nowhere to write them and the status bar says so.

## Settings

| Key                      | Purpose                                                                 |
| ------------------------ | ----------------------------------------------------------------------- |
| `anime.autoOpenJimaku`   | Pause new episodes and open Jimaku for Japanese subtitles.              |
| `anime.repos`            | Repository index URLs. Empty by default.                                |
| `anime.extensionsDir`    | Where APKs are read from. Empty uses `<userData>/anime-extensions`.     |
| `anime.preferredQuality` | Preferred stream label, matched as a substring (for example `1080`).    |
| `anime.bridgeDir`        | A bridge bundle to run instead of the downloaded one. Empty by default. |

Enable `anime.autoOpenJimaku` to hand each newly loaded Anime Browser episode
to Jimaku. SubMiner pauses playback, closes the in-player browser if it is open,
and opens Jimaku with the source title, season, and episode already filled in.
Playback resumes after the selected subtitle loads. Closing Jimaku also releases
the automatic pause, while playback that was already paused stays paused.

## Source settings

Most extensions need configuration before they return anything — a server
address and credentials, a preferred quality, a language filter. Open the
**Source settings** tab to edit them. Changes save as you make them and
persist across restarts in `<userData>/anime-source-preferences.json`.

Each save is handed back to the extension, so it can react: the Jellyfin source
logs in when the address and password land, then fills in its media-library
picker. Password-like fields are masked. Because that file can hold
credentials, it is written with owner-only permissions. Values are scoped to
the exact extension package and source, so two extensions that reuse the same
internal source ID cannot read each other's settings. Preferences saved by an
older build without package ownership are discarded; re-enter those source
settings once after upgrading.

## The bridge

SubMiner looks for the server in three places, in order, and runs the first
one it finds:

1. `anime.bridgeDir`, if set. It must hold a Java runtime and the
   `MExtensionServer-*.jar`, laid out as upstream ships them. A directory that
   holds neither is an error rather than a silent fallback.
2. A package-manager install. On Arch that is the AUR
   [`mangatan-extension-server`](https://aur.archlinux.org/packages/mangatan-extension-server)
   package, which Mangatan also uses, at `/usr/share/mangatan/extension_server`.
   Installing it means no download, and pacman keeps the bridge current. The
   `subminer-bin` package lists it as an optional dependency.
3. SubMiner's own copy in `<userData>/anime-bridge`. The first launch downloads
   the newest upstream release that ships a bundle for your platform (~130 MB,
   containing the server and a matching Java runtime, so no system JDK is
   required), the same way Mangatan does. Releases older than the oldest server
   SubMiner is known to work with are skipped. Progress appears in the banner
   at the top of the window.

The **Bridge** note at the bottom of the Extensions tab says which one is in
use, its version, and who updates it.

### Updating the bridge

Only the copy SubMiner downloaded is ever updated by SubMiner. A package-manager
install or an `anime.bridgeDir` belongs to whoever put it there.

SubMiner records the release it installed and, once the bridge is running,
asks GitHub for the newest one. When upstream has published a newer release,
the banner reads "Extension bridge v… is installed; v… is available" with an
**Update to v…** button. Clicking it downloads the new release beside the
running bridge, so a failed download changes nothing, then stops the bridge,
swaps the directories, and starts it again. The restart takes a few seconds
and kills the stream of an episode that is playing, the same as when the
bridge exits for any other reason; the queue and browser state survive. The
check is one unauthenticated GitHub API call per bridge start; if it fails
(offline, rate limited) it is logged and no update is offered until the next
start.

Upstream publishes no checksums, so the download is trusted the way Mangatan
and the AUR package trust it: TLS to GitHub and the maintainer's account. That
is the same trust running the server implies in the first place.

The bridge stays running while the window is open. Resolved video URLs point at
its own loopback proxy so the extension's cookies and headers apply, which means
those URLs stop working once it exits — the window keeps it alive for the whole
session.

If the bridge dies anyway (killed by hand, crashed, or stopped mid-operation),
the exit is detected and named in the status bar, and the next request starts a
new one. Playback already in flight still ends when its stream URL dies, but the
browser recovers without an app restart.

Two known limits:

- There is no Android WebView, so extensions that need one (typically for
  Cloudflare challenges) will fail with an error from the source.
- Bundles are published for macOS (arm64, x64), Linux (x64), and Windows (x64).
  Other platforms are unsupported outright.

## Playback

Selecting an episode resolves the best available stream, applies the source's
required headers as mpv `http-header-fields`, and loads it. The headers are
readable back off mpv, so Anki card audio and screenshots fetch correctly too.

HLS streams are routed through a small local proxy before mpv sees them. Some
hosts disguise their video segments by prepending a fake image header (a real
1x1 PNG) so scrapers back off; Aniyomi's own player strips this, but ffmpeg
probes the segment as a picture and playback dies with "no audio or video data
played". The proxy scans each segment for the first genuine MPEG-TS packet run
and drops whatever junk sits in front of it. Segments that are not TS (fMP4,
subtitles, encryption keys) pass through untouched, and direct-file streams
skip the proxy entirely. Segment URLs disguised behind fake extensions
(`.image`, `.jpg`, `.css`, and friends) are exposed locally with a `.ts`
suffix so current ffmpeg releases accept them when Anki extracts audio,
screenshots, or animated images from the playing stream.

"Playing" in the status bar means playing: after handing mpv the stream,
SubMiner waits until mpv actually configures a video output before reporting
success. If mpv gives up instead — a dead host, an undecodable stream — the
browser shows mpv's error rather than pretending playback started (a failed
load leaves no mpv window, because the player idles windowless).

Choosing **Queue** resolves the episode and appends the playable stream to mpv's
own playlist immediately, while its subtitle tracks cache in the background.
That makes mpv's next command available at once and lets the next episode begin
automatically when the current one ends, without waiting for another source
request. Resolved stream URLs can be short-lived, so a very long queue can still
outlive what its source issued; dequeue and queue that episode again to refresh
it.

### Japanese audio, and switching tracks

Sources often return a dub and the original audio as two separate entries — or
as two audio tracks of one stream — and the dub is frequently listed first.
SubMiner always aims at the Japanese audio:

- Entries labelled as a dub are skipped as long as another entry exists. This
  outranks `anime.preferredQuality`: a 1080p dub is the wrong file, not a better
  one. If every entry is a dub, it still plays.
- mpv's `alang` is set to `ja,jpn,jp,japanese` before the file loads, so a
  stream carrying several audio tracks starts on the Japanese one. With no
  Japanese track, mpv falls back to the first one as usual.
- Any audio or subtitle tracks the extension supplies separately are added to
  mpv with `audio-add` / `sub-add`, tagged with their language, and the
  Japanese one is selected.

The primary subtitle slot is reserved for Japanese — it is what the overlay
mines. A source that only carries, say, English subtitles does not get them
promoted to primary; instead the track is added with a normalized language tag
(`English` → `en`), and the regular [dual-subtitle settings](configuration.md)
apply: with `secondarySub.autoLoadSecondarySub` enabled and the language listed
in `secondarySub.secondarySubLanguages`, it is picked up as the secondary
subtitle, exactly as it would be for a local file.

Every track is added, including the ones that are not selected, so all of them
appear in mpv's track menu and can be switched by hand while watching
(`#` cycles audio, `j` cycles subtitles by default).

The extension hands over subtitle tracks as URLs, but SubMiner downloads each
one to a temporary directory and gives mpv the local file. mpv is happy either
way; [Subsync](/troubleshooting#subtitle-sync-subsync) is not, because alass
needs a file on disk to use as the timing reference. A track that fails to
download falls back to its URL so the episode still plays, the format is
detected from the file's own content rather than its URL, and the directory is
removed when the next episode starts or the app exits.

### Series, season, and episode

The episode's identity travels with it instead of being guessed back out of the
stream URL, which carries nothing but a proxy path and a file extension. The
title and episode label from the source's own listing are split into series,
season, and episode number once, at launch, and everything downstream reads
those fields:

- mpv's title reads `Series S03E04 - Episode Name`.
- Stats group by series, and rewatching an episode reuses its entry instead of
  creating a new one.
- The [Jimaku](/jimaku-integration) and [TsukiHime](/tsukihime-integration)
  modals open with Title, Season, and Episode already filled in, so a subtitle
  search is one keypress rather than a retype.
- [AniList](/anilist-integration) progress updates use those fields directly.

[mes]: https://github.com/1Selxo/M-Extension-Server
