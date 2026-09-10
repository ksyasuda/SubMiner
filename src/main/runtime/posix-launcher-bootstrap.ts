export const MANAGED_LAUNCHER_MARKER = 'SubMiner managed launcher (bundled runtime)';

export function shellQuote(value: string): string {
  return `'${value.replaceAll("'", "'\\''")}'`;
}

// Only a missing or stale Linux cache starts Electron in Node mode. Normal
// launches do one stat and execute the cached Bun and matching CLI directly.
export function posixLauncherBootstrapContent(appPath = ''): string {
  return `#!/bin/sh
# ${MANAGED_LAUNCHER_MARKER}
export SUBMINER_MANAGED_LAUNCHER=1
export SUBMINER_LAUNCHER_PATH="$0"
subminer_default_app=${shellQuote(appPath)}
case "\${XDG_DATA_HOME:-}" in
  /*) subminer_data="$XDG_DATA_HOME" ;;
  *) subminer_data="$HOME/.local/share" ;;
esac
subminer_cache="$subminer_data/SubMiner/launcher"
subminer_saved_app=
if [ -f "$subminer_cache/app-path" ]; then
  IFS= read -r subminer_saved_app < "$subminer_cache/app-path" || :
fi
subminer_app=
for subminer_candidate in "\${SUBMINER_APPIMAGE_PATH:-}" "\${SUBMINER_BINARY_PATH:-}" "$subminer_default_app" "$subminer_saved_app" "$HOME/.local/bin/SubMiner.AppImage" /opt/SubMiner/SubMiner.AppImage /Applications/SubMiner.app/Contents/MacOS/SubMiner "$HOME/Applications/SubMiner.app/Contents/MacOS/SubMiner"; do
  if [ -n "$subminer_candidate" ] && [ -x "$subminer_candidate" ]; then
    subminer_app="$subminer_candidate"
    break
  fi
done
if [ -z "$subminer_app" ]; then
  echo 'SubMiner app not found. Install the app or set SUBMINER_BINARY_PATH to its executable.' >&2
  exit 1
fi
export SUBMINER_BINARY_PATH="$subminer_app"
case "$subminer_app" in
  */Contents/MacOS/*)
    subminer_resources="\${subminer_app%/MacOS/*}/Resources"
    if [ ! -x "$subminer_resources/bun/bun" ] || [ ! -f "$subminer_resources/launcher/subminer.js" ]; then
      echo 'This launcher requires a SubMiner app with the included Bun runtime. Update SubMiner.' >&2
      exit 1
    fi
    exec "$subminer_resources/bun/bun" "$subminer_resources/launcher/subminer.js" "$@"
    ;;
esac
subminer_fingerprint=$(PATH="/usr/bin:/bin:$PATH" stat -Lc '%d:%i:%s:%y:%z' -- "$subminer_app") || exit 1
subminer_cached_fingerprint=
if [ -f "$subminer_cache/fingerprint" ]; then
  IFS= read -r subminer_cached_fingerprint < "$subminer_cache/fingerprint" || :
fi
if [ "$subminer_app" != "$subminer_saved_app" ] || [ "$subminer_fingerprint" != "$subminer_cached_fingerprint" ] || [ ! -x "$subminer_cache/bun" ] || [ ! -f "$subminer_cache/subminer" ]; then
  PATH="/usr/bin:/bin:$PATH" ELECTRON_RUN_AS_NODE=1 "$subminer_app" -e 'const p=require("node:path"); const r=process.env.APPDIR ? p.join(process.env.APPDIR,"resources") : p.join(p.dirname(process.execPath),"resources"); try { require(p.join(r,"launcher/prepare.cjs")).prepareLauncherRuntime({appPath:process.env.SUBMINER_BINARY_PATH,resourcesPath:r}); } catch(e) { console.error("Cannot prepare SubMiner launcher. Update or reinstall the SubMiner app.",e.message); process.exit(1); }' || exit $?
fi
exec "$subminer_cache/bun" "$subminer_cache/subminer" "$@"
`;
}
