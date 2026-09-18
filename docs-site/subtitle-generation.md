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

With speech detection configured, SubMiner keeps detected speech and other audible sections for Whisper to evaluate. A low speech score alone does not discard audio, which helps retain dialogue mixed with music. Only confidently silent gaps outside detected speech are omitted, with extra audio retained around each passage to reduce clipped syllables.

Passages that fit within Whisper's 30-second audio window stay intact. Longer passages prefer nearby detected speech starts when choosing cuts, falling back to quiet pauses, with a small overlap to provide context. This reduces early subtitles caused by starting a clip well before its dialogue, while keeping all retained audio covered. Matching overlapping cues are combined even when punctuation differs; repeated dialogue at separate times remains separate. SubMiner runs each passage in a fresh Whisper process so decoder state from earlier audio cannot affect later passages. This reloads the model for each passage and can increase generation time. Subtitle cues stay within the supplied audio and retain each passage's position on the original timeline. Progress reports the current passage.

This mode favors retaining dialogue over excluding music, so songs and background sounds may also produce subtitles. It can take longer than transcribing only VAD-approved speech. Whisper can still miss or misrecognize dialogue, and its timestamps remain estimates. Uncheck **Focus on spoken dialogue** to disable VAD for the session, or clear `vadModelPath` to change the default. Without a usable subtitle reference, disabling VAD returns to full-audio transcription. A selected detector or model that fails stops generation with an error. Existing subtitles are preserved.

## Using loaded subtitles as timing references

When generating for the video currently open in mpv, SubMiner automatically looks for a dialogue subtitle track among its embedded subtitles and loaded external SRT, ASS/SSA, or WebVTT files. It prefers English, then tracks labeled full or dialogue. Forced tracks, image subtitles, generated subtitles, and tracks whose titles or filenames identify signs, songs, lyrics, karaoke, or opening/ending subtitles are skipped. These checks rely on metadata; an unlabeled signs-only file cannot always be identified.

The selected reference appears in generation progress. SubMiner reads its timestamps, including the active primary or secondary track's subtitle delay, and uses nearby cue starts to guide cuts in long audio passages. Short passages stay intact. The reference works with or without **Focus on spoken dialogue**. With that option enabled, reference starts take priority over VAD starts when choosing a nearby cut; VAD still helps identify speech. Audio outside reference cues remains eligible for transcription, and Whisper still supplies the Japanese text and final timestamps. The reference is assumed to be timed for the playing video; this does not automatically sync a mistimed reference.

Unreadable or empty references are skipped in favor of another eligible loaded track. If none can be read, generation uses its normal audio timing. The launcher uses loaded references only when its input matches the video currently open in mpv; standalone generation keeps its existing behavior. It captures reference tracks and delays together with the initial audio selection, before checking or downloading a model. When relying on mpv's selected audio, it stops and asks you to retry if the media changes or cannot be verified during capture.

## Choosing a model

The modal recommends **large-v3-turbo** when it detects an NVIDIA GPU through `nvidia-smi` and the selected `whisper-cli` discovers an available CUDA device. Otherwise it recommends **small** for a balance of Japanese recognition quality and CPU time. The check works before downloading a model and falls back to small if a tool is missing, fails, times out, or reports an unrecognized result. Vulkan, AMD, and Apple GPU support do not qualify for the turbo recommendation. Checks are cached for up to 30 seconds; changing the Whisper executable path triggers a new check.

The recommendation labels the model picker and explains the detected support. It does not change your configured model, current selection, external model path, or launcher's model choice. It also does not force a GPU backend during transcription. These are starting recommendations, not hardware benchmarks or guarantees that every model fits in available GPU memory. Tiny and base need less memory and usually finish sooner, with more recognition errors. Medium and large models favor accuracy but need more resources. Large-v3-turbo is optimized for speed compared with large-v3, with some accuracy tradeoff; actual performance depends on your CPU, GPU, whisper.cpp build, and audio.

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
