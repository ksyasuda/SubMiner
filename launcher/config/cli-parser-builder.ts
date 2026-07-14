import { Command } from 'commander';

export interface JellyfinInvocation {
  action?: string;
  discovery?: boolean;
  play?: boolean;
  login?: boolean;
  logout?: boolean;
  setup?: boolean;
  server?: string;
  username?: string;
  password?: string;
  passwordStore?: string;
  logLevel?: string;
}

export interface CommandActionInvocation {
  action?: string;
  logLevel?: string;
}

export interface CliInvocations {
  jellyfinInvocation: JellyfinInvocation | null;
  configInvocation: CommandActionInvocation | null;
  settingsInvocation: CommandActionInvocation | null;
  mpvInvocation: CommandActionInvocation | null;
  appInvocation: { appArgs: string[] } | null;
  dictionaryTriggered: boolean;
  dictionaryTarget: string | null;
  dictionaryLogLevel: string | null;
  dictionaryCandidates: boolean;
  dictionarySelect: boolean;
  dictionaryAnilistId: string | null;
  statsTriggered: boolean;
  statsBackground: boolean;
  statsStop: boolean;
  statsCleanup: boolean;
  statsCleanupVocab: boolean;
  statsCleanupLifetime: boolean;
  statsLogLevel: string | null;
  syncTriggered: boolean;
  syncCliTokens: string[];
  syncLogLevel: string | null;
  syncUiTriggered: boolean;
  syncUiLogLevel: string | null;
  doctorTriggered: boolean;
  doctorLogLevel: string | null;
  doctorRefreshKnownWords: boolean;
  logsTriggered: boolean;
  logsExport: boolean;
  texthookerTriggered: boolean;
  texthookerLogLevel: string | null;
  texthookerOpenBrowser: boolean;
}

function applyRootOptions(program: Command): void {
  program
    .option(
      '-b, --backend <backend>',
      'Display backend (auto, hyprland, sway, x11, macos, windows)',
    )
    .option('-d, --directory <dir>', 'Directory to browse')
    .option('-a, --args <args>', 'Pass arguments to MPV')
    .option('-r, --recursive', 'Search directories recursively')
    .option('-p, --profile <profile>', 'MPV profile')
    .option('--start', 'Explicitly start overlay')
    .option('--log-level <level>', 'Log level')
    .option('-v, --version', 'Show SubMiner version')
    .option('--settings', 'Open settings window')
    .option('-u, --update', 'Check for updates')
    .option('-R, --rofi', 'Use rofi picker')
    .option('-H, --history', 'Browse local watch history')
    .option('-S, --start-overlay', 'Auto-start overlay')
    .option('-T, --no-texthooker', 'Disable texthooker-ui server')
    // The SubMiner app answers sync commands when invoked with --sync-cli.
    // Remote-command resolution may address the launcher the same way, so
    // accept the flag as a no-op to keep both invocation shapes equivalent.
    .option('--sync-cli', 'Compatibility no-op (sync commands work with or without it)');
}

function buildSubcommandHelpText(program: Command): string {
  const subcommands = program.commands
    .filter((command) => command.name() !== 'help')
    .map((command) => {
      const aliases = command.aliases();
      const term = aliases.length > 0 ? `${command.name()}|${aliases[0]}` : command.name();
      return { term, description: command.description() };
    });

  if (subcommands.length === 0) return '';
  const longestTerm = Math.max(...subcommands.map((entry) => entry.term.length));
  const lines = subcommands.map((entry) =>
    `  ${entry.term.padEnd(longestTerm)}  ${entry.description || ''}`.trimEnd(),
  );
  return `\nCommands:\n${lines.join('\n')}\n`;
}

function getTopLevelCommand(argv: string[]): { name: string; index: number } | null {
  const commandNames = new Set([
    'jellyfin',
    'jf',
    'doctor',
    'config',
    'settings',
    'mpv',
    'logs',
    'dictionary',
    'dict',
    'stats',
    'sync',
    'texthooker',
    'app',
    'bin',
    'help',
  ]);
  const optionsWithValue = new Set([
    '-b',
    '--backend',
    '-a',
    '--args',
    '-d',
    '--directory',
    '-p',
    '--profile',
    '--log-level',
  ]);
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i] || '';
    if (token === '--') return null;
    if (token.startsWith('-')) {
      if (optionsWithValue.has(token)) i += 1;
      continue;
    }
    return commandNames.has(token) ? { name: token, index: i } : null;
  }
  return null;
}

function hasTopLevelCommand(argv: string[]): boolean {
  return getTopLevelCommand(argv) !== null;
}

export function resolveTopLevelCommand(argv: string[]): { name: string; index: number } | null {
  return getTopLevelCommand(argv);
}

export function parseCliPrograms(
  argv: string[],
  scriptName: string,
): {
  options: Record<string, unknown>;
  rootTarget: unknown;
  invocations: CliInvocations;
} {
  let jellyfinInvocation: JellyfinInvocation | null = null;
  let configInvocation: CommandActionInvocation | null = null;
  let settingsInvocation: CommandActionInvocation | null = null;
  let mpvInvocation: CommandActionInvocation | null = null;
  let appInvocation: { appArgs: string[] } | null = null;
  let dictionaryTriggered = false;
  let dictionaryTarget: string | null = null;
  let dictionaryLogLevel: string | null = null;
  let dictionaryCandidates = false;
  let dictionarySelect = false;
  let dictionaryAnilistId: string | null = null;
  let statsTriggered = false;
  let statsBackground = false;
  let statsStop = false;
  let statsCleanup = false;
  let statsCleanupVocab = false;
  let statsCleanupLifetime = false;
  let statsLogLevel: string | null = null;
  let syncTriggered = false;
  let syncCliTokens: string[] = [];
  let syncLogLevel: string | null = null;
  let syncUiTriggered = false;
  let syncUiLogLevel: string | null = null;
  let doctorLogLevel: string | null = null;
  let doctorRefreshKnownWords = false;
  let logsTriggered = false;
  let logsExport = false;
  let texthookerLogLevel: string | null = null;
  let texthookerOpenBrowser = false;
  let doctorTriggered = false;
  let texthookerTriggered = false;

  const commandProgram = new Command();
  commandProgram
    .name(scriptName)
    .description('Launch MPV with SubMiner sentence mining overlay')
    .showHelpAfterError(true)
    .enablePositionalOptions()
    .allowExcessArguments(false)
    .allowUnknownOption(false)
    .exitOverride();
  applyRootOptions(commandProgram);

  const rootProgram = new Command();
  rootProgram
    .name(scriptName)
    .description('Launch MPV with SubMiner sentence mining overlay')
    .usage('[options] [command] [target]')
    .showHelpAfterError(true)
    .allowExcessArguments(false)
    .allowUnknownOption(false)
    .exitOverride()
    .argument('[target]', 'file, directory, or URL');
  applyRootOptions(rootProgram);

  commandProgram
    .command('jellyfin')
    .alias('jf')
    .description('Jellyfin workflows')
    .argument('[action]', 'setup|discovery|play|login|logout')
    .option('-d, --discovery', 'Cast discovery mode')
    .option('-p, --play', 'Interactive play picker')
    .option('-l, --login', 'Login flow')
    .option('--logout', 'Clear token/session')
    .option('--setup', 'Open setup window')
    .option('-s, --server <url>', 'Jellyfin server URL')
    .option('-u, --username <name>', 'Jellyfin username')
    .option('-w, --password <pass>', 'Jellyfin password')
    .option('--password-store <backend>', 'Pass through Electron safeStorage backend')
    .option('--log-level <level>', 'Log level')
    .action((action: string | undefined, options: Record<string, unknown>) => {
      jellyfinInvocation = {
        action,
        discovery: options.discovery === true,
        play: options.play === true,
        login: options.login === true,
        logout: options.logout === true,
        setup: options.setup === true,
        server: typeof options.server === 'string' ? options.server : undefined,
        username: typeof options.username === 'string' ? options.username : undefined,
        password: typeof options.password === 'string' ? options.password : undefined,
        passwordStore:
          typeof options.passwordStore === 'string' ? options.passwordStore : undefined,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('dictionary')
    .alias('dict')
    .description('Generate or correct character dictionary AniList matches')
    .argument('[target]', 'Video file path or anime directory path')
    .option('--candidates', 'List AniList candidates for a character dictionary target')
    .option('--select <id>', 'Pin an AniList media ID for the target series')
    .option('--log-level <level>', 'Log level')
    .action((target: string | undefined, options: Record<string, unknown>) => {
      const selectValue = typeof options.select === 'string' ? options.select.trim() : '';
      const hasSelect = selectValue.length > 0;
      if (options.candidates === true && hasSelect) {
        throw new Error('Dictionary --candidates and --select cannot be combined.');
      }
      dictionaryTriggered = true;
      dictionaryTarget = target ?? null;
      dictionaryLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
      dictionaryCandidates = options.candidates === true;
      dictionarySelect = hasSelect;
      dictionaryAnilistId = hasSelect ? selectValue : null;
    });

  commandProgram
    .command('stats')
    .description('Launch the local immersion stats dashboard')
    .argument('[action]', 'cleanup|rebuild|backfill')
    .option('-b, --background', 'Start the stats server in the background')
    .option('-s, --stop', 'Stop the background stats server')
    .option('-v, --vocab', 'Clean vocabulary rows in the stats database')
    .option('-l, --lifetime', 'Rebuild lifetime summary rows from retained data')
    .option('--log-level <level>', 'Log level')
    .action((action: string | undefined, options: Record<string, unknown>) => {
      statsTriggered = true;
      const normalizedAction = (action || '').toLowerCase();
      statsBackground = options.background === true;
      statsStop = options.stop === true;
      if (statsBackground && statsStop) {
        throw new Error('Stats background and stop flags cannot be combined.');
      }
      if (
        normalizedAction &&
        normalizedAction !== 'cleanup' &&
        normalizedAction !== 'rebuild' &&
        normalizedAction !== 'backfill'
      ) {
        throw new Error('Invalid stats action. Valid values are cleanup, rebuild, or backfill.');
      }
      if (normalizedAction && (statsBackground || statsStop)) {
        throw new Error('Stats background and stop flags cannot be combined with stats actions.');
      }
      if (normalizedAction !== 'cleanup' && (options.vocab === true || options.lifetime === true)) {
        throw new Error('Stats --vocab and --lifetime flags require the cleanup action.');
      }
      if (normalizedAction === 'cleanup') {
        statsCleanup = true;
        statsCleanupLifetime = options.lifetime === true;
        statsCleanupVocab = statsCleanupLifetime ? false : options.vocab !== false;
      } else if (normalizedAction === 'rebuild' || normalizedAction === 'backfill') {
        statsCleanup = true;
        statsCleanupLifetime = true;
        statsCleanupVocab = false;
      }
      statsLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
    });

  commandProgram
    .command('sync')
    .description('Sync stats and watch history with another machine over SSH')
    .argument('[host]', 'SSH destination (user@host or an ssh config alias)')
    .option('--snapshot <file>', 'Write a consistent snapshot of the local stats database')
    .option('--merge <file>', 'Merge a snapshot database file into the local stats database')
    .option('--push', 'Only merge local stats into the SSH host')
    .option('--pull', 'Only merge stats from the SSH host into the local database')
    .option('--db <file>', 'Override the local stats database path')
    .option('--remote-cmd <cmd>', 'subminer command to run on the remote host')
    .option('-f, --force', 'Skip the running-app safety check')
    .option('--check', 'Test the SSH connection and remote subminer availability')
    .option('--json', 'Emit machine-readable NDJSON progress output')
    .option('--make-temp', 'Create a sync temp directory and print its path (used over SSH)')
    .option('--remove-temp <dir>', 'Remove a sync temp directory created by --make-temp')
    .option('--ui', 'Open the SubMiner sync window')
    .option('--log-level <level>', 'Log level')
    .action((rawHost: string | undefined, options: Record<string, unknown>) => {
      const host = typeof rawHost === 'string' ? rawHost.trim() : '';
      const snapshot = typeof options.snapshot === 'string' ? options.snapshot.trim() : '';
      const merge = typeof options.merge === 'string' ? options.merge.trim() : '';
      const push = options.push === true;
      const pull = options.pull === true;
      const check = options.check === true;
      const makeTemp = options.makeTemp === true;
      const removeTemp = typeof options.removeTemp === 'string' ? options.removeTemp.trim() : '';
      if (options.ui === true) {
        if (
          host ||
          snapshot ||
          merge ||
          push ||
          pull ||
          check ||
          makeTemp ||
          removeTemp ||
          options.remoteCmd !== undefined ||
          options.db !== undefined ||
          options.json === true ||
          options.force === true
        ) {
          throw new Error('Sync --ui cannot be combined with other sync options.');
        }
        syncUiTriggered = true;
        syncUiLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
        return;
      }
      // No validation here: the app's parseSyncCliTokens owns the sync rules
      // and its error text reaches the terminal through the child's stdio.
      const remoteCmd = typeof options.remoteCmd === 'string' ? options.remoteCmd.trim() : '';
      const dbPath = typeof options.db === 'string' ? options.db.trim() : '';
      const tokens: string[] = [];
      if (host) tokens.push(host);
      if (snapshot) tokens.push('--snapshot', snapshot);
      if (merge) tokens.push('--merge', merge);
      if (makeTemp) tokens.push('--make-temp');
      if (removeTemp) tokens.push('--remove-temp', removeTemp);
      if (push) tokens.push('--push');
      if (pull) tokens.push('--pull');
      if (check) tokens.push('--check');
      if (remoteCmd) tokens.push('--remote-cmd', remoteCmd);
      if (dbPath) tokens.push('--db', dbPath);
      if (options.force === true) tokens.push('--force');
      if (options.json === true) tokens.push('--json');
      syncTriggered = true;
      syncCliTokens = tokens;
      syncLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
    });

  commandProgram
    .command('doctor')
    .description('Run dependency and environment checks')
    .option('--refresh-known-words', 'Refresh known words cache')
    .option('--log-level <level>', 'Log level')
    .action((options: Record<string, unknown>) => {
      doctorTriggered = true;
      doctorLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
      doctorRefreshKnownWords = options.refreshKnownWords === true;
    });

  commandProgram
    .command('logs')
    .description('Log file helpers')
    .option('-e, --export', 'Export sanitized log archive')
    .action((options: Record<string, unknown>) => {
      logsTriggered = true;
      logsExport = options.export === true;
    });

  commandProgram
    .command('config')
    .description('Config file helpers (path|show)')
    .argument('[action]', 'path|show')
    .option('--log-level <level>', 'Log level')
    .action((action: string | undefined, options: Record<string, unknown>) => {
      configInvocation = {
        action,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('settings')
    .description('Open SubMiner settings window')
    .option('--log-level <level>', 'Log level')
    .action((options: Record<string, unknown>) => {
      settingsInvocation = {
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('mpv')
    .description('MPV helpers')
    .argument('[action]', 'status|socket|idle', 'status')
    .option('--log-level <level>', 'Log level')
    .action((action: string, options: Record<string, unknown>) => {
      mpvInvocation = {
        action,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('texthooker')
    .description('Launch texthooker-only mode')
    .option('-o, --open-browser', 'Open texthooker in the default browser')
    .option('--log-level <level>', 'Log level')
    .action((options: Record<string, unknown>) => {
      texthookerTriggered = true;
      texthookerLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
      texthookerOpenBrowser = options.openBrowser === true;
    });

  commandProgram
    .command('app')
    .alias('bin')
    .description('Pass arguments directly to SubMiner binary')
    .allowUnknownOption(true)
    .allowExcessArguments(true)
    .argument('[appArgs...]', 'Arguments forwarded to SubMiner app binary')
    .action((appArgs: string[] | undefined) => {
      appInvocation = { appArgs: Array.isArray(appArgs) ? appArgs : [] };
    });

  rootProgram.addHelpText('after', buildSubcommandHelpText(commandProgram));

  const selectedProgram = hasTopLevelCommand(argv) ? commandProgram : rootProgram;
  try {
    selectedProgram.parse(['node', scriptName, ...argv]);
  } catch (error) {
    const commanderError = error as { code?: string; message?: string };
    if (commanderError?.code === 'commander.helpDisplayed') {
      process.exit(0);
    }
    throw new Error(commanderError?.message || String(error));
  }

  return {
    options: selectedProgram.opts<Record<string, unknown>>(),
    rootTarget: rootProgram.processedArgs[0],
    invocations: {
      jellyfinInvocation,
      configInvocation,
      settingsInvocation,
      mpvInvocation,
      appInvocation,
      dictionaryTriggered,
      dictionaryTarget,
      dictionaryLogLevel,
      dictionaryCandidates,
      dictionarySelect,
      dictionaryAnilistId,
      statsTriggered,
      statsBackground,
      statsStop,
      statsCleanup,
      statsCleanupVocab,
      statsCleanupLifetime,
      statsLogLevel,
      syncTriggered,
      syncCliTokens,
      syncLogLevel,
      syncUiTriggered,
      syncUiLogLevel,
      doctorTriggered,
      doctorLogLevel,
      doctorRefreshKnownWords,
      logsTriggered,
      logsExport,
      texthookerTriggered,
      texthookerLogLevel,
      texthookerOpenBrowser,
    },
  };
}
