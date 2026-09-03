type: fixed
area: dictionary

- Upgraded the desktop runtime to Electron 43.4.1 and added profile guards that block unsupported runtimes and Electron downgrades before Yomitan storage is loaded.
- Development launches now use a separate `SubMiner-dev` profile unless production-profile access is explicitly requested.
- Automatic character-dictionary changes now stop when a previously non-empty Yomitan profile suddenly reports zero dictionaries.
