# SubMiner vs Memento: Comparative Analysis

## Overview

|                    | **SubMiner**                          | **Memento**                              |
| ------------------ | ------------------------------------- | ---------------------------------------- |
| **Language**       | TypeScript (Electron)                 | C++ (Qt6)                                |
| **Video Player**   | External mpv (IPC socket)             | Embedded libmpv (OpenGL widget)          |
| **Dictionary**     | Yomitan (embedded browser extension)  | Own SQLite engine (Yomichan-format dicts) |
| **Platforms**      | Linux (X11/Wayland), macOS            | Linux, macOS, Windows                    |
| **Maturity**       | Early (v0.1.0)                        | Mature (v1.7.2, AUR/Flathub packages)   |

---

## Where Memento is Stronger

### 1. Self-Contained Native Application

Memento embeds mpv directly via libmpv, creating a seamless single-window experience. Player controls, subtitle display, and dictionary lookups all live in one cohesive application. SubMiner runs as a separate overlay process communicating with mpv via IPC, which introduces latency and compositor-specific complexity (Hyprland, Sway, X11 each need their own window-tracking backend).

### 2. Built-In Dictionary Engine

Memento has its own SQLite-backed dictionary engine that directly imports Yomichan zip files. No external browser extension needed. This gives Memento tighter control over the lookup UX — results appear inline, kanji cards show stroke order, pitch accent diagrams render natively. SubMiner delegates all dictionary rendering to Yomitan, so the lookup experience depends entirely on the external extension.

### 3. Deconjugation Engine

Memento has a full 90+ rule deconjugation engine as a fallback when MeCab isn't available. This means grammar-aware search works even without optional dependencies. SubMiner relies entirely on MeCab for morphological analysis and falls back to raw, unsegmented text if MeCab is unavailable.

### 4. Windows Support

Memento has fully documented Windows builds (MSYS2/MinGW), portable installs, and platform-specific troubleshooting docs. SubMiner doesn't ship Windows builds and the platform detection code is untested on Windows.

### 5. Kanji Detail Cards

Memento displays rich kanji detail cards with stroke order visualization, JLPT level, frequency rankings, classifications, and codepoints — all rendered natively in Qt. SubMiner delegates kanji display entirely to Yomitan.

### 6. Packaging & Distribution

Memento is available on AUR, Flathub, and provides pre-built binaries for all three platforms. It's an established project with a known user base and mature release process.

### 7. OCR Support

Memento supports optional OCR via manga-ocr for extracting text from on-screen images (useful for hardsubbed content or manga). SubMiner has no OCR capability.

### 8. Resource Efficiency

As a native C++/Qt application, Memento has a significantly smaller memory footprint than SubMiner's Electron-based architecture, which typically consumes 150-300MB of RAM for the Chromium runtime alone.

---

## Where SubMiner is Stronger

### 1. Anki Integration Depth

SubMiner's Anki workflow is significantly more advanced:

- **Automatic polling** detects new cards and auto-populates media fields (no manual trigger needed after the initial mine action)
- **Kiku field grouping** detects duplicate words across cards and merges them into grouped cards with a two-step preview modal
- **Dual note type support** (Lapis sentence cards + Kiku vocabulary cards) with independent field mappings
- **AI translation fallback** uses OpenAI-compatible APIs (OpenRouter, local models) to generate translations when secondary subtitles are unavailable

Memento's Anki integration is capable but more traditional: manual "Add to Anki" buttons with configurable templates and no automation beyond the initial card creation.

### 2. Subtitle Timing Intelligence

SubMiner maintains a subtitle timing cache (200-entry LRU with 5-minute TTL) and uses fuzzy edit-distance matching to locate audio segments for media extraction. This means audio clips are precisely timed even when card creation is delayed or when the subtitle text has been slightly modified. Memento captures the current frame and audio at the moment of card creation.

### 3. Multi-Copy & Batch Mining

SubMiner supports multi-copy mode (`Ctrl+Shift+C` to batch-copy N recent subtitle lines) and multi-mine mode (`Ctrl+Alt+S`). This is useful for dialogue-heavy scenes where you want context spanning multiple subtitle lines in a single card.

### 4. YouTube Subtitle Generation

SubMiner has built-in Whisper integration for generating subtitles from YouTube videos that lack them. Three modes are available: automatic (generate on-the-fly), preprocess (generate before playback), and off. Memento uses yt-dlp for streaming but has no subtitle generation capability for videos without existing subtitles.

### 5. Jimaku Integration

SubMiner connects to the Jimaku anime subtitle API for searching and downloading subtitle files directly from the overlay UI. This streamlines the workflow of finding Japanese subtitles for anime. Memento has no equivalent feature.

### 6. Subtitle Sync (Subsync)

SubMiner integrates with alass and ffsubsync for in-session subtitle resynchronization, with a modal UI for selecting the sync engine and source audio track. Memento relies on mpv's native subtitle delay controls, which only handle constant offsets rather than non-linear drift.

### 7. Texthooker / WebSocket Broadcasting

SubMiner runs a texthooker HTTP server (port 5174) and a WebSocket server (port 6677), broadcasting subtitles and tokenization data to external tools in real-time. This enables integration with browser-based dictionaries, other mining tools, or custom automation workflows. Memento is fully self-contained with no external broadcast mechanism.

### 8. AI Translation

SubMiner can call OpenAI-compatible APIs to generate translations on-the-fly when secondary subtitles aren't available. This includes configurable system prompts, target language selection, and auto-retry with exponential backoff. Memento has no AI or machine translation capability.

### 9. Runtime Options Modal

SubMiner has a `Ctrl+Shift+O` palette for toggling session-only options (auto-update cards, field grouping, etc.) without editing config files or restarting. Memento requires navigating through its settings dialog for configuration changes.

### 10. Overlay Architecture

SubMiner's transparent overlay design means it works with any existing mpv instance — users keep their mpv configuration, plugins, shaders, and keybindings untouched. Memento requires using its own player, and while it supports mpv config files, they must be placed in Memento's separate config directory.

---

## Where They're Comparable

| Feature                  | SubMiner                              | Memento                               |
| ------------------------ | ------------------------------------- | ------------------------------------- |
| MeCab tokenization       | Yes (optional)                        | Yes (optional)                        |
| Secondary subtitles      | Yes (3 display modes)                 | Yes (requires mpv >= 0.35)            |
| Configurable styling     | CSS variables in config               | Qt stylesheet + settings dialog       |
| mpv config support       | Via mpv directly (user's own config)  | Separate Memento config directory     |
| AnkiConnect protocol     | Yes                                   | Yes                                   |
| Media capture            | FFmpeg (audio + image/AVIF)           | mpv screenshot + audio extraction     |
| Yomichan dict compat     | Via Yomitan extension                 | Native import of Yomichan zip files   |
| Pitch accent display     | Via Yomitan                           | Native rendering                      |
| Streaming (yt-dlp)       | Yes                                   | Yes                                   |

---

## Architectural Trade-offs

### Overlay vs Integrated App

| Aspect          | SubMiner (Overlay)                                           | Memento (Integrated)                                    |
| --------------- | ------------------------------------------------------------ | ------------------------------------------------------- |
| UX cohesion     | Separate window; requires compositor cooperation             | Single window; everything in one place                  |
| mpv flexibility | Works with any mpv instance and config                       | Must use Memento's player; separate config directory    |
| Latency         | IPC socket introduces small delay                            | Direct libmpv calls; near-zero latency                  |
| Platform quirks | Needs per-compositor window tracking (Hyprland/Sway/X11)     | Qt handles platform abstraction                         |

### Electron vs Native C++/Qt

| Aspect             | SubMiner (Electron/TypeScript)              | Memento (C++/Qt6)                         |
| ------------------ | ------------------------------------------- | ----------------------------------------- |
| Dev velocity       | Rapid iteration, rich npm ecosystem         | Slower development, manual memory mgmt    |
| Memory footprint   | ~150-300MB (Chromium runtime)               | ~30-80MB typical                          |
| Startup time       | Slower (Chromium spin-up)                   | Fast native startup                       |
| Dependency weight   | Node modules + Electron binary (~200MB)     | System libraries (Qt, mpv, SQLite)        |
| UI flexibility     | Full HTML/CSS/JS; easy to style             | Qt widgets; powerful but more constrained |

### External Yomitan vs Built-In Dictionary

| Aspect             | SubMiner (Yomitan)                                | Memento (Built-in)                            |
| ------------------ | ------------------------------------------------- | --------------------------------------------- |
| Maintenance burden | Free updates from Yomitan project                 | Must maintain dictionary engine in-house      |
| UX control         | Limited to Yomitan's popup behavior               | Full control over lookup UI and rendering     |
| Setup complexity   | Requires Yomitan extension + dictionaries         | Import dictionaries directly; no extension    |
| Feature parity     | Gets all Yomitan features automatically           | Must implement each feature (pitch, kanji)    |

---

## Code Quality Comparison

### SubMiner

**Strengths:**
- Clean TypeScript with strict mode
- Centralized config registry with validation and example generation
- Robust error handling with exponential backoff retry logic
- Good separation of concerns (tokenization, media generation, Anki polling)

**Weaknesses:**
- Monolithic `main.ts` at ~5,000 lines — could benefit from modularization
- Limited test coverage (only `config.test.ts` exists)
- Token merging heuristics are complex (14 helper functions) and may miss edge cases

### Memento

**Strengths:**
- Modular C++ architecture with clear directory structure (dict/, player/, gui/, anki/)
- Qt signal/slot pattern keeps components decoupled
- Thread-safe dictionary access via `QReadWriteLock`
- Smart pointer usage throughout; no obvious memory management issues
- Settings migration system for version upgrades

**Weaknesses:**
- 90+ hardcoded deconjugation rules need manual maintenance for new patterns
- No visible test suite in the repository
- MpvWidget property map handles 40+ properties in a single class

---

## Summary

**Memento** is a more mature, polished, self-contained desktop application. Its strength is the integrated experience — one app that handles video playback, dictionary lookups, and Anki card creation in a single native window. It has better Windows support, a built-in dictionary engine with deconjugation fallback, kanji stroke-order cards, and OCR support. It's the safer choice for a user who wants something that works out of the box with minimal setup.

**SubMiner** is more innovative in its Anki workflow automation, mining UX, and external integrations. Features like automatic card polling, Kiku field grouping, AI translation fallback, Whisper subtitle generation, Jimaku subtitle search, subsync integration, and WebSocket broadcasting are all capabilities Memento doesn't have. The overlay architecture also means users can keep their existing mpv setup completely untouched.

### Key Memento Advantages to Learn From

- Built-in deconjugation fallback (grammar awareness without MeCab)
- OCR support for non-subtitle text
- Windows platform support
- Self-contained dictionary engine (no external browser extension dependency)
- Lower resource footprint as a native application

### Key SubMiner Advantages to Protect

- Anki automation depth (polling, field grouping, AI translation)
- Subtitle intelligence (timing cache, Whisper generation, Jimaku, subsync)
- External tool interop (WebSocket, texthooker server)
- Overlay model (works with any mpv instance and configuration)
- Rapid feature development enabled by TypeScript/Electron
