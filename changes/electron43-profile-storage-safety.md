type: fixed
area: dictionary

- Upgraded the desktop runtime to Electron 43.7.2 and added profile guards that block unsupported runtimes and Electron major or minor downgrades before Yomitan storage is loaded. Patch downgrades within the same major and minor version are allowed.
- Development launches now use a separate `SubMiner-dev` profile unless production-profile access is explicitly requested.
- Automatic character-dictionary changes now stop when a previously non-empty Yomitan profile suddenly reports zero dictionaries.
