# Contributor Note

To add or change a config option, update `src/config/definitions.ts` first. Defaults, runtime-option metadata, and generated `config.example.jsonc` are derived from this centralized source.

## Environment Variables

| Variable                 | Description                                    |
| ------------------------ | ---------------------------------------------- |
| `SUBMINER_APPIMAGE_PATH` | Override AppImage location for subminer script |
| `SUBMINER_YT_SUBGEN_MODE` | Override `youtubeSubgen.mode` for launcher |
| `SUBMINER_WHISPER_BIN` | Override `youtubeSubgen.whisperBin` for launcher |
| `SUBMINER_WHISPER_MODEL` | Override `youtubeSubgen.whisperModel` for launcher |
| `SUBMINER_YT_SUBGEN_OUT_DIR` | Override generated subtitle output directory |
| `SUBMINER_YT_SUBGEN_AUDIO_FORMAT` | Override extraction format used for whisper fallback |
| `SUBMINER_YT_SUBGEN_KEEP_TEMP` | Set to `1` to keep temporary subtitle-generation workspace |
