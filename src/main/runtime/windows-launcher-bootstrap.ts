const MANAGED_LAUNCHER_MARKER = 'SubMiner managed launcher (bundled runtime)';

function windowsBatchLiteral(value: string): string {
  if (/["\r\n]/.test(value)) {
    throw new Error('Launcher paths cannot contain quotes or newlines.');
  }
  return value.replaceAll('%', '%%');
}

function configuredAppCandidate(appPath: string | undefined): string[] {
  if (!appPath) return [];
  const literal = windowsBatchLiteral(appPath);
  return [
    `if not exist "${literal}" goto subminer_check_local_app`,
    `set "SUBMINER_BINARY_PATH=${literal}"`,
    'goto subminer_app_found',
  ];
}

// This command file stays valid across app updates. It locates the current app
// and only starts Electron in Node mode when that version's private Bun is absent.
export function windowsLauncherBootstrapContent(appPath?: string): string {
  return [
    '@echo off',
    `rem ${MANAGED_LAUNCHER_MARKER}`,
    'setlocal DisableDelayedExpansion',
    'set "SUBMINER_MANAGED_LAUNCHER=1"',
    'set "SUBMINER_LAUNCHER_PATH=%~f0"',
    'if defined SUBMINER_BINARY_PATH goto subminer_app_found',
    ...configuredAppCandidate(appPath),
    ':subminer_check_local_app',
    'if not defined LOCALAPPDATA goto subminer_check_program_files',
    'if not exist "%LOCALAPPDATA%\\Programs\\SubMiner\\SubMiner.exe" goto subminer_check_program_files',
    'set "SUBMINER_BINARY_PATH=%LOCALAPPDATA%\\Programs\\SubMiner\\SubMiner.exe"',
    'goto subminer_app_found',
    ':subminer_check_program_files',
    'if not defined ProgramFiles goto subminer_app_missing',
    'if not exist "%ProgramFiles%\\SubMiner\\SubMiner.exe" goto subminer_app_missing',
    'set "SUBMINER_BINARY_PATH=%ProgramFiles%\\SubMiner\\SubMiner.exe"',
    ':subminer_app_found',
    'if not exist "%SUBMINER_BINARY_PATH%" goto subminer_app_missing',
    'if not defined LOCALAPPDATA goto subminer_local_app_data_missing',
    'for %%I in ("%SUBMINER_BINARY_PATH%") do set "SUBMINER_RESOURCES_PATH=%%~dpIresources"',
    'if not exist "%SUBMINER_RESOURCES_PATH%\\launcher\\subminer.js" goto subminer_resources_missing',
    'set "SUBMINER_APP_VERSION="',
    'if not exist "%SUBMINER_RESOURCES_PATH%\\launcher\\version" goto subminer_resources_missing',
    'set /p "SUBMINER_APP_VERSION="<"%SUBMINER_RESOURCES_PATH%\\launcher\\version"',
    'if not defined SUBMINER_APP_VERSION goto subminer_resources_missing',
    'set "SUBMINER_BUN_PATH=%LOCALAPPDATA%\\SubMiner\\launcher-runtime\\%SUBMINER_APP_VERSION%\\bun.exe"',
    'if exist "%SUBMINER_BUN_PATH%" goto subminer_run',
    'if not exist "%SUBMINER_RESOURCES_PATH%\\launcher\\prepare.cjs" goto subminer_resources_missing',
    'set "ELECTRON_RUN_AS_NODE=1"',
    "\"%SUBMINER_BINARY_PATH%\" -e \"const p=require('node:path');try{require(p.join(process.env.SUBMINER_RESOURCES_PATH,'launcher','prepare.cjs')).prepareLauncherRuntime({appPath:process.env.SUBMINER_BINARY_PATH,resourcesPath:process.env.SUBMINER_RESOURCES_PATH});}catch(error){console.error('Cannot prepare SubMiner launcher. Update or reinstall the SubMiner app.',error instanceof Error?error.message:String(error));process.exit(1);}\"",
    'set "SUBMINER_PREPARE_EXIT=%errorlevel%"',
    'set "ELECTRON_RUN_AS_NODE="',
    'if not "%SUBMINER_PREPARE_EXIT%"=="0" exit /b %SUBMINER_PREPARE_EXIT%',
    'if not exist "%SUBMINER_BUN_PATH%" goto subminer_prepare_missing',
    ':subminer_run',
    '"%SUBMINER_BUN_PATH%" "%SUBMINER_RESOURCES_PATH%\\launcher\\subminer.js" %*',
    'exit /b %errorlevel%',
    ':subminer_app_missing',
    '>&2 echo SubMiner app not found. Install the app or set SUBMINER_BINARY_PATH to its executable.',
    'exit /b 1',
    ':subminer_local_app_data_missing',
    '>&2 echo LOCALAPPDATA is unavailable. SubMiner cannot locate its private launcher runtime.',
    'exit /b 1',
    ':subminer_resources_missing',
    '>&2 echo This launcher requires a SubMiner app with the included Bun runtime. Update SubMiner.',
    'exit /b 1',
    ':subminer_prepare_missing',
    '>&2 echo SubMiner did not create its private launcher runtime. Update or reinstall SubMiner.',
    'exit /b 1',
    '',
  ].join('\r\n');
}
