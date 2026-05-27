type: added
area: setup

- Added optional first-run setup controls to install Bun and the `subminer` command-line launcher on Linux, macOS, and Windows, with a Windows `subminer.cmd` PATH shim so `subminer` works without manually adding `SubMiner.exe` to PATH.
- Added an Open SubMiner Settings button to first-run setup and moved Finish to the right-side action slot.
- First-run setup recognizes existing `subminer` installs in Homebrew or user PATH directories, while manual setup avoids writing into Homebrew-owned paths.
- The standalone setup app quits after completing first-run setup, returning the terminal instead of leaving the process open.
