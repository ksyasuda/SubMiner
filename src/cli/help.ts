export function printHelp(defaultTexthookerPort: number): void {
  const tty = process.stdout?.isTTY ?? false;
  const B = tty ? '\x1b[1m' : '';
  const D = tty ? '\x1b[2m' : '';
  const R = tty ? '\x1b[0m' : '';

  console.log(`
${B}SubMiner${R} — Japanese sentence mining with mpv + Yomitan

${B}Usage:${R} subminer ${D}[command] [options]${R}

${B}Session${R}
  --background                   Start in tray/background mode
  --start                        Connect to mpv and launch overlay
  --launch-mpv ${D}[targets...]${R}       Launch mpv with SubMiner defaults and exit
  --stop                         Stop the running instance
  --stats                        Open the stats dashboard in your browser
  --texthooker                   Start texthooker server only ${D}(no overlay)${R}
  --open-browser                 Open texthooker in your default browser
  --update                       Check for updates

${B}Overlay${R}
  --toggle-visible-overlay       Toggle subtitle overlay
  --toggle-primary-subtitle-bar  Toggle primary subtitle bar
  --show-visible-overlay         Show subtitle overlay
  --hide-visible-overlay         Hide subtitle overlay
  --yomitan                      Open Yomitan settings window
  --settings                     Open SubMiner settings window
  --setup                        Open first-run setup window
  --auto-start-overlay           Auto-hide mpv subs, show overlay on connect

${B}Mining${R}
  --mine-sentence                Create Anki card from current subtitle
  --mine-sentence-multiple       Select multiple lines, then mine
  --copy-subtitle                Copy current subtitle to clipboard
  --copy-subtitle-multiple       Enter multi-line copy mode
  --update-last-card-from-clipboard  Update last Anki card from clipboard
  --mark-audio-card              Mark last card as audio-only
  --trigger-field-grouping       Run Kiku field grouping
  --trigger-subsync              Run subtitle sync
  --toggle-secondary-sub         Cycle secondary subtitle mode
  --mark-watched                 Mark current video watched and advance playlist
  --toggle-subtitle-sidebar      Toggle subtitle sidebar panel
  --open-runtime-options         Open runtime options palette
  --open-session-help            Open session help modal
  --open-character-dictionary    Open character dictionary anime selection modal
  --open-controller-select       Open controller select modal
  --open-controller-debug        Open controller debug modal

${B}AniList${R}
  --anilist-setup                Open AniList authentication flow
  --anilist-status               Show token and retry queue status
  --anilist-logout               Clear stored AniList token
  --anilist-retry-queue          Retry next queued update
  --dictionary                   Generate character dictionary ZIP for current anime
  --dictionary-candidates        Show character dictionary AniList candidates
  --dictionary-select            Save manual character dictionary AniList selection
  --dictionary-anilist-id ${D}ID${R}      AniList media ID for --dictionary-select
  --dictionary-target ${D}PATH${R}         Override dictionary source path (file or directory)

${B}Jellyfin${R}
  --jellyfin                     Open Jellyfin setup window
  --jellyfin-login               Authenticate and store session token
  --jellyfin-logout              Clear stored session data
  --jellyfin-libraries           List available libraries
  --jellyfin-items               List items from a library
  --jellyfin-subtitles           List subtitle tracks for an item
  --jellyfin-subtitle-urls       Print subtitle download URLs only
  --jellyfin-play                Stream an item in mpv
  --jellyfin-remote-announce     Broadcast cast-target capability

  ${D}Jellyfin options:${R}
  --jellyfin-server ${D}URL${R}          Server URL ${D}(overrides config)${R}
  --jellyfin-username ${D}NAME${R}       Username for login
  --jellyfin-password ${D}PASS${R}       Password for login
  --jellyfin-library-id ${D}ID${R}       Library to browse
  --jellyfin-item-id ${D}ID${R}          Item to play or inspect
  --jellyfin-search ${D}QUERY${R}        Filter items by search term
  --jellyfin-limit ${D}N${R}             Max items returned
  --jellyfin-audio-stream-index ${D}N${R}       Audio stream override
  --jellyfin-subtitle-stream-index ${D}N${R}    Subtitle stream override

${B}Options${R}
  --socket ${D}PATH${R}                  mpv IPC socket path
  --backend ${D}BACKEND${R}              Window tracker ${D}(auto, hyprland, sway, x11, macos, windows)${R}
  --port ${D}PORT${R}                    Texthooker server port ${D}(default: ${defaultTexthookerPort})${R}
  --log-level ${D}LEVEL${R}              ${D}debug | info | warn | error${R}
  --debug                        Enable debug mode ${D}(alias: --dev)${R}
  --generate-config              Write default config.jsonc
  --config-path ${D}PATH${R}             Target path for --generate-config
  --backup-overwrite             Backup existing config before overwrite
  --help                         Show this help
`);
}
