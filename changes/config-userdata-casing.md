type: fixed
area: startup

- Prevented early Electron startup from writing config and user data under a lowercase `~/.config/subminer` path instead of the canonical `~/.config/SubMiner` directory.
