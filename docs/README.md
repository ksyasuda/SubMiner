# Documentation

SubMiner documentation is built with [VitePress](https://vitepress.dev/).

## Local Docs Site

```bash
make docs-dev     # Dev server at http://localhost:5173
make docs         # Build static output
make docs-preview # Preview built site at http://localhost:4173
```

## Pages

### Getting Started

- [Installation](/installation) — Requirements, Linux/macOS/Windows install, mpv plugin setup
- [Usage](/usage) — `subminer` wrapper + subcommands (`jellyfin`, `yt`, `doctor`, `config`, `mpv`, `texthooker`, `app`), mpv plugin, keybindings
- [Mining Workflow](/mining-workflow) — End-to-end sentence mining guide, overlay layers, card creation

### Reference

- [Configuration](/configuration) — Full config file reference and option details
- [Keyboard Shortcuts](/shortcuts) — All global, overlay, mining, and plugin chord shortcuts in one place
- [Anki Integration](/anki-integration) — AnkiConnect setup, field mapping, media generation, field grouping
- [Jellyfin Integration](/jellyfin-integration) — Optional Jellyfin auth, cast discovery, remote control, and playback launch
- [Immersion Tracking](/immersion-tracking) — SQLite schema, retention/rollup policies, query templates, and extension points
- [JLPT Vocabulary](/jlpt-vocab-bundle) — Bundled term-meta bank for JLPT level underlining and frequency highlighting
- [MPV Plugin](/mpv-plugin) — Chord keybindings, subminer.conf options, script messages
- [Troubleshooting](/troubleshooting) — Common issues and solutions by category

### Development

- [Building & Testing](/development) — Build commands, test suites, contributor notes, environment variables
- [Architecture](/architecture) — Service-oriented design, composition model, renderer module layout
