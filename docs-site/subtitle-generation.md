# Japanese subtitle generation

Generate Japanese SRT subtitles from a local video's audio using [whisper.cpp](https://github.com/ggml-org/whisper.cpp). The launcher and overlay use the same local generation service. Audio stays on your computer. Model downloads require an internet connection; generation with an installed model does not.

## Setup

Install whisper.cpp's `whisper-cli` executable and FFmpeg, including `ffprobe`. SubMiner downloads models, not these executables. Leave `whisperPath`, `ffmpegPath`, and `ffprobePath` empty to find the executables on `PATH`. To use a specific installation, set a path override under **Settings → Integrations → Japanese Subtitle Generation**.

The generation modal checks for these executables under **Local tools** and keeps **Generate subtitles** disabled until every required one is found, naming the missing executable and its setting. Model downloads stay available in the meantime. After installing a tool or changing a path, click **Check again**. The launcher runs the same check before any model download. Generation also confirms the destination directory grants write and search permissions before extracting audio.

Choose one model source:

- Set `subtitleGeneration.modelPath` to an existing **multilingual whisper.cpp GGML `.bin` model**. Python Whisper checkpoints and English-only models are not suitable for Japanese transcription.
- Leave that path empty and choose a model directly in the generation modal. Each option shows its download size; the selected model has speed and accuracy guidance. The modal offers **Download model** when that model is missing. Your choice lasts for the current SubMiner session, including closing and reopening the modal. Set `subtitleGeneration.managedModel` in Settings to change the default for future sessions.

Managed models are stored in `models/whisper/` beside your SubMiner configuration file. Downloads show progress, verify the expected file size and SHA256, and publish the model only after verification. Cancelling or failing a download removes its temporary files. A configured external path always takes precedence; an unreadable path displays an error instead of silently downloading another model.

See the [generated configuration example](/config.example.jsonc) for current defaults. Changes apply to the next operation.

## Prioritizing spoken dialogue

To focus on dialogue, check the optional **Focus on spoken dialogue** box in the generation modal. If the speech detection model is missing, click **Download speech detection model** to install it. This separate download uses the same progress, cancellation, and integrity checks as Whisper downloads. Checking the box never downloads automatically, and leaving it unchecked lets you generate without the Silero model.

You also need whisper.cpp's [speech segment detector](https://github.com/ggml-org/whisper.cpp/tree/master/examples/vad-speech-segments). SubMiner downloads the model, not this executable. The detector is found as `whisper-vad-speech-segments` or, for builds from the upstream source, `vad-speech-segments` on `PATH`. Set `vadPath` in **Settings → Integrations → Japanese Subtitle Generation** for any other location. With **Focus on spoken dialogue** checked, the modal's **Local tools** check requires the detector too.

The checkbox choice lasts for the current SubMiner session, including closing and reopening the modal. To make dialogue mode your default, set `vadModelPath` in Settings to a [Silero GGML VAD model](https://huggingface.co/ggml-org/whisper-vad/tree/main). The modal downloads `ggml-silero-v6.2.0.bin` into the same `models/whisper/` directory as managed Whisper models. An existing configured VAD path takes precedence and checks the box initially. Unchecking it temporarily disables dialogue mode without changing that path. Downloading the model alone does not enable dialogue mode.

With speech detection configured, SubMiner transcribes short speech passages separately and restores each passage's position on the original audio timeline. Detection retains brief utterances and includes extra audio around speech to reduce clipped syllables. If the detector returns a long passage, SubMiner looks for quiet pauses near chunk boundaries. Adjacent chunks overlap slightly to provide context when speech continues through a cut. Matching subtitle cues in that overlap are combined; repeated dialogue at separate times remains separate.

SubMiner resets transcription context between passages and limits subtitle cues to the audio supplied for each chunk. A line cannot stretch across an omitted music break. Progress reports completed batches of dialogue passages. These adjustments do not replace Whisper's timestamp estimates or guarantee that every spoken line is recognized.

This mode prioritizes spoken dialogue over songs and background sounds. It can miss quiet speech or speech mixed with loud music, and recognition errors are still possible. Uncheck **Focus on spoken dialogue** to return to full-audio transcription for the session, or clear `vadModelPath` to change the default. A selected detector or model that fails stops generation with an error. Existing subtitles are preserved.

## Choosing a model

Start with **small** for a balance of Japanese recognition quality and CPU time. This is a general starting recommendation, not a benchmark for your hardware. Tiny and base need less memory and usually finish sooner, with more recognition errors. Medium and large models favor accuracy but need more resources. Large-v3-turbo is optimized for speed compared with large-v3, with some accuracy tradeoff; actual performance depends on your CPU, GPU, whisper.cpp build, and audio.

The picker includes whisper.cpp's official multilingual tiny, base, small, medium, large-v1, large-v2, large-v3, and large-v3-turbo downloads, including their available quantized variants. Quantized models use less disk space and memory, with possible accuracy loss. English-only `.en` models are excluded. See the [upstream model list](https://github.com/ggml-org/whisper.cpp/blob/master/models/download-ggml-model.sh) and [Whisper's model guidance](https://github.com/openai/whisper#available-models-and-languages).

A configured external Model Path takes precedence and hides the managed model picker. Clear it in Settings to choose a managed model. Changing the picker never downloads automatically, and it cannot change the model during an active download or generation.

## From the overlay

1. Open a local video in mpv and select its Japanese audio track.
2. Press **Ctrl+Shift+G** to open the standalone generation modal. When the subtitle sidebar has no subtitle lines loaded, it also offers a **Generate Japanese subtitles** button. Neither an open sidebar nor an existing subtitle track is required for the shortcut.
3. Choose a model and download it if prompted, or configure your existing model path in Settings and click **Check again**.
4. Optionally check **Focus on spoken dialogue** and click **Download speech detection model** if prompted.
5. Click **Generate subtitles**.

The modal shows audio preparation, transcription, and saving progress. Percentages appear when the underlying tool reports them. **Cancel** stops the current operation. Closing the modal lets the job continue; reopening it shows the current progress or result.

**Escape** or **Close** closes the modal using the same focus and overlay restoration as other SubMiner modals. Change or disable its shortcut with `shortcuts.openSubtitleGeneration` in Settings. Ctrl+G remains assigned to field grouping.

SubMiner saves `<video>.ja.generated.srt` beside the media, adding a numeric suffix if that name already exists. It selects the generated Japanese subtitle track and resets the subtitle delay when mpv is still playing the same file. If playback changes, the subtitles remain saved and are not attached to the new video. The result includes the saved path even if mpv cannot load it.

## From the launcher

```bash
subminer generate-subs episode.mkv --download-model
subminer generate-subs episode.mkv --model-path /path/to/ggml-small.bin
subminer generate-subs
```

With no file argument, the command uses the current local mpv media and its selected audio track. With an explicit file, it prefers an audio stream tagged Japanese, otherwise the first audio stream. Use `--audio-stream` to choose an absolute FFmpeg stream index. `--output` specifies a new destination SRT; existing output files are never overwritten. See [launcher usage](/usage) for all flags. Ctrl+C cancels the operation.

## Timing and limitations

The SRT includes whisper.cpp's timestamps, adjusted for the audio stream's position on the media timeline and, when speech detection is configured, each passage's original start time. No alass step is required to load it. This version uses native Whisper timing; it does not run WhisperX or another forced aligner. Recognition can repeat or invent lines, and timing can be imperfect, especially with music or overlapping speech. Review generated text and audio boundaries when mining.

Generation supports local files and internal audio tracks. Remote URLs, subtitle translation, and transcription of a separately attached mpv audio track are not supported by the modal. Pass a separate local audio file to the launcher if needed. The destination directory needs writable space for subtitles; temporary storage needs enough space for the extracted mono audio.
