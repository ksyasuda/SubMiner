import * as childProcess from "child_process";

const SECRET_COMMAND_PATTERN = /^\((.*)\)$/s;
const COMMAND_CACHE = new Map<string, string>();
const COMMAND_PENDING = new Map<string, Promise<string>>();

function executeCommand(command: string): Promise<string> {
  return new Promise((resolve, reject) => {
    childProcess.exec(command, { timeout: 10_000 }, (err, stdout) => {
      if (err) {
        reject(err);
        return;
      }
      resolve(stdout);
    });
  });
}

export function clearAnilistClientSecretCache(): void {
  COMMAND_CACHE.clear();
  COMMAND_PENDING.clear();
}

function resolveCommand(rawSecret: string): string | null {
  const commandMatch = rawSecret.match(SECRET_COMMAND_PATTERN);
  if (!commandMatch || commandMatch[1] === undefined) {
    return null;
  }

  const command = commandMatch[1].trim();
  return command.length > 0 ? command : null;
}

export async function resolveAnilistClientSecret(rawSecret: string): Promise<string> {
  const trimmedSecret = rawSecret.trim();
  if (trimmedSecret.length === 0) {
    throw new Error("cannot authenticate without client secret");
  }

  const command = resolveCommand(trimmedSecret);
  if (!command) {
    return trimmedSecret;
  }

  const cachedValue = COMMAND_CACHE.get(trimmedSecret);
  if (cachedValue !== undefined) {
    return cachedValue;
  }

  const pending = COMMAND_PENDING.get(trimmedSecret);
  if (pending !== undefined) {
    return pending;
  }

  const promise = executeCommand(command)
    .then((stdout) => {
      const trimmed = stdout.trim();
      if (!trimmed) {
        throw new Error("secret command returned empty value");
      }
      COMMAND_CACHE.set(trimmedSecret, trimmed);
      COMMAND_PENDING.delete(trimmedSecret);
      return trimmed;
    })
    .catch((error) => {
      COMMAND_PENDING.delete(trimmedSecret);
      const errorMessage = error instanceof Error ? error.message : String(error);
      if (errorMessage === "secret command returned empty value") {
        throw error;
      }
      throw new Error(`secret command failed: ${errorMessage}`);
    });

  COMMAND_PENDING.set(trimmedSecret, promise);
  return promise;
}
