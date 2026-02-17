import * as fs from "fs";
import * as path from "path";
import * as readline from "readline";
import { CliArgs } from "../../cli/args";
import { createLogger } from "../../logger";

const logger = createLogger("core:config-gen");

function formatBackupTimestamp(date = new Date()): string {
  const pad = (v: number): string => String(v).padStart(2, "0");
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function promptYesNo(question: string): Promise<boolean> {
  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question(question, (answer) => {
      rl.close();
      const normalized = answer.trim().toLowerCase();
      resolve(normalized === "y" || normalized === "yes");
    });
  });
}

export async function generateDefaultConfigFile(
  args: CliArgs,
  options: {
    configDir: string;
    defaultConfig: unknown;
    generateTemplate: (config: unknown) => string;
  },
): Promise<number> {
  const targetPath = args.configPath
    ? path.resolve(args.configPath)
    : path.join(options.configDir, "config.jsonc");
  const template = options.generateTemplate(options.defaultConfig);

  if (fs.existsSync(targetPath)) {
    if (args.backupOverwrite) {
      const backupPath = `${targetPath}.bak.${formatBackupTimestamp()}`;
      fs.copyFileSync(targetPath, backupPath);
      fs.writeFileSync(targetPath, template, "utf-8");
      logger.info(`Backed up existing config to ${backupPath}`);
      logger.info(`Generated config at ${targetPath}`);
      return 0;
    }

    if (!process.stdin.isTTY || !process.stdout.isTTY) {
      logger.error(
        `Config exists at ${targetPath}. Re-run with --backup-overwrite to back up and overwrite.`,
      );
      return 1;
    }

    const confirmed = await promptYesNo(
      `Config exists at ${targetPath}. Back up and overwrite? [y/N] `,
    );
    if (!confirmed) {
      logger.info("Config generation cancelled.");
      return 0;
    }

    const backupPath = `${targetPath}.bak.${formatBackupTimestamp()}`;
    fs.copyFileSync(targetPath, backupPath);
    fs.writeFileSync(targetPath, template, "utf-8");
    logger.info(`Backed up existing config to ${backupPath}`);
    logger.info(`Generated config at ${targetPath}`);
    return 0;
  }

  const parentDir = path.dirname(targetPath);
  if (!fs.existsSync(parentDir)) {
    fs.mkdirSync(parentDir, { recursive: true });
  }
  fs.writeFileSync(targetPath, template, "utf-8");
  logger.info(`Generated config at ${targetPath}`);
  return 0;
}
