export function printHelp(defaultTexthookerPort: number): void {
  console.log(`
SubMiner CLI commands:
  --start               Start MPV IPC connection and overlay control loop
  --stop                Stop the running overlay app
  --toggle              Toggle visible subtitle overlay visibility (legacy alias)
  --toggle-visible-overlay    Toggle visible subtitle overlay visibility
  --toggle-invisible-overlay  Toggle invisible interactive overlay visibility
  --settings            Open Yomitan settings window
  --texthooker          Launch texthooker only (no overlay window)
  --show                Force show visible overlay (legacy alias)
  --hide                Force hide visible overlay (legacy alias)
  --show-visible-overlay       Force show visible subtitle overlay
  --hide-visible-overlay       Force hide visible subtitle overlay
  --show-invisible-overlay     Force show invisible interactive overlay
  --hide-invisible-overlay     Force hide invisible interactive overlay
  --copy-subtitle              Copy current subtitle text
  --copy-subtitle-multiple     Start multi-copy mode
  --mine-sentence              Mine sentence card from current subtitle
  --mine-sentence-multiple     Start multi-mine sentence mode
   --update-last-card-from-clipboard  Update last card from clipboard
   --refresh-known-words          Refresh known words cache now
   --toggle-secondary-sub       Cycle secondary subtitle mode
  --trigger-field-grouping     Trigger Kiku field grouping
  --trigger-subsync            Run subtitle sync
  --mark-audio-card            Mark last card as audio card
  --open-runtime-options       Open runtime options palette
  --auto-start-overlay  Auto-hide mpv subtitles on connect (show overlay)
  --socket PATH         Override MPV IPC socket/pipe path
  --backend BACKEND     Override window tracker backend (auto, hyprland, sway, x11, macos)
  --port PORT           Texthooker server port (default: ${defaultTexthookerPort})
  --verbose             Enable debug logging (equivalent to --log-level debug)
  --log-level LEVEL     Set log level: debug, info, warn, error
  --generate-config     Generate default config.jsonc from centralized config registry
  --config-path PATH    Target config path for --generate-config
  --backup-overwrite    With --generate-config, backup and overwrite existing file
  --dev                 Run in development mode
  --debug               Alias for --dev
  --help                Show this help
`);
}
