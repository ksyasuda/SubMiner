type: fixed
area: release

Linux packaged builds now expose the canonical `SubMiner` app identity to Electron's startup metadata so native Wayland compositors stop reporting the window class/app-id as lowercase `subminer`.
