import { BrowserWindow, Extension, session } from "electron";
import * as fs from "fs";
import * as path from "path";

export interface YomitanExtensionLoaderDeps {
  userDataPath: string;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  setYomitanExtension: (extension: Extension | null) => void;
}

function ensureExtensionCopy(sourceDir: string, userDataPath: string): string {
  if (process.platform === "win32") {
    return sourceDir;
  }

  const extensionsRoot = path.join(userDataPath, "extensions");
  const targetDir = path.join(extensionsRoot, "yomitan");

  const sourceManifest = path.join(sourceDir, "manifest.json");
  const targetManifest = path.join(targetDir, "manifest.json");

  let shouldCopy = !fs.existsSync(targetDir);
  if (
    !shouldCopy &&
    fs.existsSync(sourceManifest) &&
    fs.existsSync(targetManifest)
  ) {
    try {
      const sourceVersion = (
        JSON.parse(fs.readFileSync(sourceManifest, "utf-8")) as {
          version: string;
        }
      ).version;
      const targetVersion = (
        JSON.parse(fs.readFileSync(targetManifest, "utf-8")) as {
          version: string;
        }
      ).version;
      shouldCopy = sourceVersion !== targetVersion;
    } catch {
      shouldCopy = true;
    }
  }

  if (shouldCopy) {
    fs.mkdirSync(extensionsRoot, { recursive: true });
    fs.rmSync(targetDir, { recursive: true, force: true });
    fs.cpSync(sourceDir, targetDir, { recursive: true });
    console.log(`Copied yomitan extension to ${targetDir}`);
  }

  return targetDir;
}

export async function loadYomitanExtensionService(
  deps: YomitanExtensionLoaderDeps,
): Promise<Extension | null> {
  const searchPaths = [
    path.join(__dirname, "..", "..", "vendor", "yomitan"),
    path.join(__dirname, "..", "..", "..", "vendor", "yomitan"),
    path.join(process.resourcesPath, "yomitan"),
    "/usr/share/SubMiner/yomitan",
    path.join(deps.userDataPath, "yomitan"),
  ];

  let extPath: string | null = null;
  for (const p of searchPaths) {
    if (fs.existsSync(p)) {
      extPath = p;
      break;
    }
  }

  if (!extPath) {
    console.error("Yomitan extension not found in any search path");
    console.error("Install Yomitan to one of:", searchPaths);
    return null;
  }

  extPath = ensureExtensionCopy(extPath, deps.userDataPath);

  const parserWindow = deps.getYomitanParserWindow();
  if (parserWindow && !parserWindow.isDestroyed()) {
    parserWindow.destroy();
  }
  deps.setYomitanParserWindow(null);
  deps.setYomitanParserReadyPromise(null);
  deps.setYomitanParserInitPromise(null);

  try {
    const extensions = session.defaultSession.extensions;
    const extension = extensions
      ? await extensions.loadExtension(extPath, {
          allowFileAccess: true,
        })
      : await session.defaultSession.loadExtension(extPath, {
          allowFileAccess: true,
        });
    deps.setYomitanExtension(extension);
    return extension;
  } catch (err) {
    console.error("Failed to load Yomitan extension:", (err as Error).message);
    console.error("Full error:", err);
    deps.setYomitanExtension(null);
    return null;
  }
}
