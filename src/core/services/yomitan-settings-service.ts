import { BrowserWindow, Extension, session } from "electron";

export function openYomitanSettingsWindow(options: {
  yomitanExt: Extension | null;
  getExistingWindow: () => BrowserWindow | null;
  setWindow: (window: BrowserWindow | null) => void;
}): void {
  console.log("openYomitanSettings called");

  if (!options.yomitanExt) {
    console.error(
      "Yomitan extension not loaded - yomitanExt is:",
      options.yomitanExt,
    );
    console.error(
      "This may be due to Manifest V3 service worker issues with Electron",
    );
    return;
  }

  const existingWindow = options.getExistingWindow();
  if (existingWindow && !existingWindow.isDestroyed()) {
    console.log("Settings window already exists, focusing");
    existingWindow.focus();
    return;
  }

  console.log("Creating new settings window for extension:", options.yomitanExt.id);

  const settingsWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      session: session.defaultSession,
    },
  });
  options.setWindow(settingsWindow);

  const settingsUrl = `chrome-extension://${options.yomitanExt.id}/settings.html`;
  console.log("Loading settings URL:", settingsUrl);

  let loadAttempts = 0;
  const maxAttempts = 3;

  const attemptLoad = (): void => {
    settingsWindow
      .loadURL(settingsUrl)
      .then(() => {
        console.log("Settings URL loaded successfully");
      })
      .catch((err: Error) => {
        console.error("Failed to load settings URL:", err);
        loadAttempts++;
        if (loadAttempts < maxAttempts && !settingsWindow.isDestroyed()) {
          console.log(
            `Retrying in 500ms (attempt ${loadAttempts + 1}/${maxAttempts})`,
          );
          setTimeout(attemptLoad, 500);
        }
      });
  };

  attemptLoad();

  settingsWindow.webContents.on(
    "did-fail-load",
    (_event, errorCode, errorDescription) => {
      console.error(
        "Settings page failed to load:",
        errorCode,
        errorDescription,
      );
    },
  );

  settingsWindow.webContents.on("did-finish-load", () => {
    console.log("Settings page loaded successfully");
  });

  setTimeout(() => {
    if (!settingsWindow.isDestroyed()) {
      settingsWindow.setSize(
        settingsWindow.getSize()[0],
        settingsWindow.getSize()[1],
      );
      settingsWindow.webContents.invalidate();
      settingsWindow.show();
    }
  }, 500);

  settingsWindow.on("closed", () => {
    options.setWindow(null);
  });
}
