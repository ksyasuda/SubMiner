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
- [Usage](/usage) — `subminer` wrapper, mpv plugin, keybindings, YouTube playback
- [Mining Workflow](/mining-workflow) — End-to-end sentence mining guide, overlay layers, card creation

### Reference

- [Configuration](/configuration) — Full config file reference and option details
- [Anki Integration](/anki-integration) — AnkiConnect setup, field mapping, media generation, field grouping
- [MPV Plugin](/mpv-plugin) — Chord keybindings, subminer.conf options, script messages
- [Troubleshooting](/troubleshooting) — Common issues and solutions by category

### Development

- [Building & Testing](/development) — Build commands, test suites, contributor notes, environment variables
- [Architecture](/architecture) — Service-oriented design, composition model, renderer module layout
