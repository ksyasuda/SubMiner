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

export interface YtInvocation {
  target?: string;
  outDir?: string;
  keepTemp?: boolean;
  whisperBin?: string;
  whisperModel?: string;
  whisperVadModel?: string;
  whisperThreads?: number;
  ytSubgenAudioFormat?: string;
  logLevel?: string;
}

export interface CommandActionInvocation {
  action: string;
  logLevel?: string;
}

export interface CliInvocations {
  jellyfinInvocation: JellyfinInvocation | null;
  ytInvocation: YtInvocation | null;
  configInvocation: CommandActionInvocation | null;
  mpvInvocation: CommandActionInvocation | null;
  appInvocation: { appArgs: string[] } | null;
  dictionaryTriggered: boolean;
  dictionaryTarget: string | null;
  dictionaryLogLevel: string | null;
  statsTriggered: boolean;
  statsBackground: boolean;
  statsStop: boolean;
  statsCleanup: boolean;
  statsCleanupVocab: boolean;
  statsCleanupLifetime: boolean;
  statsLogLevel: string | null;
  doctorTriggered: boolean;
  doctorLogLevel: string | null;
  texthookerTriggered: boolean;
  texthookerLogLevel: string | null;
}

function applyRootOptions(program: Command): void {
  program
    .option('-b, --backend <backend>', 'Display backend')
    .option('-d, --directory <dir>', 'Directory to browse')
    .option('-a, --args <args>', 'Pass arguments to MPV')
    .option('-r, --recursive', 'Search directories recursively')
    .option('-p, --profile <profile>', 'MPV profile')
    .option('--start', 'Explicitly start overlay')
    .option('--log-level <level>', 'Log level')
    .option('-R, --rofi', 'Use rofi picker')
    .option('-S, --start-overlay', 'Auto-start overlay')
    .option('-T, --no-texthooker', 'Disable texthooker-ui server');
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
    'yt',
    'youtube',
    'doctor',
    'config',
    'mpv',
    'dictionary',
    'dict',
    'stats',
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
  let ytInvocation: YtInvocation | null = null;
  let configInvocation: CommandActionInvocation | null = null;
  let mpvInvocation: CommandActionInvocation | null = null;
  let appInvocation: { appArgs: string[] } | null = null;
  let dictionaryTriggered = false;
  let dictionaryTarget: string | null = null;
  let dictionaryLogLevel: string | null = null;
  let statsTriggered = false;
  let statsBackground = false;
  let statsStop = false;
  let statsCleanup = false;
  let statsCleanupVocab = false;
  let statsCleanupLifetime = false;
  let statsLogLevel: string | null = null;
  let doctorLogLevel: string | null = null;
  let texthookerLogLevel: string | null = null;
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
    .command('yt')
    .alias('youtube')
    .description('YouTube workflows')
    .argument('[target]', 'YouTube URL or ytsearch: query')
    .option('-o, --out-dir <dir>', 'Subtitle output dir')
    .option('--keep-temp', 'Keep temp files')
    .option('--whisper-bin <path>', 'whisper.cpp CLI path')
    .option('--whisper-model <path>', 'whisper model path')
    .option('--whisper-vad-model <path>', 'whisper.cpp VAD model path')
    .option('--whisper-threads <n>', 'whisper.cpp thread count')
    .option('--yt-subgen-audio-format <format>', 'Audio extraction format')
    .option('--log-level <level>', 'Log level')
    .action((target: string | undefined, options: Record<string, unknown>) => {
      ytInvocation = {
        target,
        outDir: typeof options.outDir === 'string' ? options.outDir : undefined,
        keepTemp: options.keepTemp === true,
        whisperBin: typeof options.whisperBin === 'string' ? options.whisperBin : undefined,
        whisperModel: typeof options.whisperModel === 'string' ? options.whisperModel : undefined,
        whisperVadModel:
          typeof options.whisperVadModel === 'string' ? options.whisperVadModel : undefined,
        whisperThreads:
          typeof options.whisperThreads === 'number' && Number.isFinite(options.whisperThreads)
            ? Math.floor(options.whisperThreads)
            : undefined,
        ytSubgenAudioFormat:
          typeof options.ytSubgenAudioFormat === 'string' ? options.ytSubgenAudioFormat : undefined,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('dictionary')
    .alias('dict')
    .description('Generate character dictionary ZIP from a file or directory target')
    .argument('<target>', 'Video file path or anime directory path')
    .option('--log-level <level>', 'Log level')
    .action((target: string, options: Record<string, unknown>) => {
      dictionaryTriggered = true;
      dictionaryTarget = target;
      dictionaryLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
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
        throw new Error(
          'Invalid stats action. Valid values are cleanup, rebuild, or backfill.',
        );
      }
      if (normalizedAction && (statsBackground || statsStop)) {
        throw new Error('Stats background and stop flags cannot be combined with stats actions.');
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
    .command('doctor')
    .description('Run dependency and environment checks')
    .option('--log-level <level>', 'Log level')
    .action((options: Record<string, unknown>) => {
      doctorTriggered = true;
      doctorLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
    });

  commandProgram
    .command('config')
    .description('Config helpers')
    .argument('[action]', 'path|show', 'path')
    .option('--log-level <level>', 'Log level')
    .action((action: string, options: Record<string, unknown>) => {
      configInvocation = {
        action,
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
    .option('--log-level <level>', 'Log level')
    .action((options: Record<string, unknown>) => {
      texthookerTriggered = true;
      texthookerLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
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
      ytInvocation,
      configInvocation,
      mpvInvocation,
      appInvocation,
      dictionaryTriggered,
      dictionaryTarget,
      dictionaryLogLevel,
      statsTriggered,
      statsBackground,
      statsStop,
      statsCleanup,
      statsCleanupVocab,
      statsCleanupLifetime,
      statsLogLevel,
      doctorTriggered,
      doctorLogLevel,
      texthookerTriggered,
      texthookerLogLevel,
    },
  };
}
