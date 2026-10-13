type: docs
area: launcher

- Removed `SUBMINER_JIMAKU_API_KEY`, `SUBMINER_JIMAKU_API_KEY_COMMAND`, and `SUBMINER_JIMAKU_API_BASE_URL` from the development docs' environment variable table. The launcher never used them; set the Jimaku key with `jimaku.apiKey` / `jimaku.apiKeyCommand` in the config instead.
