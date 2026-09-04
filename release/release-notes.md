## Highlights
### Added
- **Pre-Mining Timing Review**:
  - Optional review step before creating word, sentence, or audio cards, with a speech-focused waveform that filters out steady background noise so dialogue is easy to spot.
  - The clip end automatically snaps back to where dialogue actually ends, since subtitles often linger after speech stops.
  - Drag or use the keyboard to adjust clip boundaries, and preview audio with a sweeping playhead that plays to the true end even on high-latency outputs like Bluetooth headphones.
  - Pull extra previous or next subtitle lines onto the card with `P`/`N` (or the Prev/Next steppers); a live preview shows exactly what the card will contain.
  - You can cancel and still keep the card without media, and the review can be toggled on or off for the session.
- **Senren Note Type Support**:
  - Enable `ankiConnect.isSenren` to merge duplicate mined cards using Senren's scene-switching markup, combining sentence, furigana, audio, picture, and misc-info fields.
  - Supports the same auto/manual/disabled grouping modes as Kiku, including the manual merge modal. Senren and Kiku are mutually exclusive, so only one can be enabled at a time.

### Changed
- **Remote Streaming Mining**: Mining a card from a remote stream (Jellyfin and other HTTP sources) now downloads the clip window once and reuses it for the timing review waveform, audio preview, audio extraction, and screenshot, instead of re-fetching the stream for every step. No action needed; the temporary download is cleaned up automatically after ten minutes of inactivity.
- **TsukiHime Release Picker**: The Japanese and secondary-language tabs now filter releases down to ones that actually carry subtitles for that language, and tell you when none do.

### Fixed
- **Broadcast Caption Accuracy**:
  - Japanese caption tracks split across two positioned lines (e.g. Crunchyroll) now merge into one, so mined sentences, the sidebar, and line-break settings treat them as a single line; lines from different speakers or sound effects still stay separate.
  - Mining from the overlay no longer picks up a leftover line from the previous caption, so the mined sentence and clip timing match what's actually on screen.
  - Copying or mining subtitles no longer includes the separate furigana line that some broadcast subtitle files place above kanji.
- **Multi-line Copy After Seeking**: Selecting multiple subtitle lines to copy or mine now selects backward in timeline order after a seek, rather than in playback encounter order.
- **Overlay Stability**:
  - On Hyprland, opening a modal (timing review, Jimaku, session help, and others) over fullscreen mpv no longer makes the overlay flicker while the modal loads.
  - Switching secondary subtitle tracks no longer causes mpv's native secondary subtitles to flash on screen.
- **Anki Update Notifications**: Switching notification settings to on-screen display while a card update is still in progress now correctly dismisses the old overlay progress indicator.
- **Jellyfin Subtitles**: Subtitle files now load with zero delay in mpv instead of Jellyfin inferring and applying a sync offset.

## What's Changed

- feat(anki): add media timing review before card creation by @ksyasuda in #203
- fix(jellyfin): stop inferring subtitle delays by @ksyasuda in #227
- feat(anki): support Senren scene-switching field grouping by @ksyasuda in #230
- fix(mining): copy multi-line subtitles backward from current line by @ksyasuda in #231
- fix(subtitles): keep native secondary subtitles hidden by @ksyasuda in #232
- fix(subtitles): drop ASS furigana from recorded cues by @ksyasuda in #233
- fix(subtitles): merge wrapped positioned caption rows by @ksyasuda in #234
- fix(tsukihime): filter releases by subtitle language by @ksyasuda in #235

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
