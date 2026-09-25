# Japanese subtitle generation

When a video has no Japanese subtitles, SubMiner can transcribe its audio into a Japanese SRT with [whisper.cpp](https://github.com/ggml-org/whisper.cpp). Everything runs on your computer. You only need internet access to download a model.

## Setup

1. Install whisper.cpp's `whisper-cli` and FFmpeg (including `ffprobe`). SubMiner downloads models but not these programs.
2. Make sure they are on your `PATH`, or set their paths under **Settings > Integrations > Japanese Subtitle Generation** (`whisperPath`, `ffmpegPath`, `ffprobePath`).
3. Pick a model. Either choose one in the generation modal and click **Download model**, or set `subtitleGeneration.modelPath` to a multilingual whisper.cpp GGML `.bin` file you already have. English-only models and Python Whisper checkpoints do not work.

Downloaded models go to `models/whisper/` next to your SubMiner config file. A configured `modelPath` always wins over the modal's choice.

The modal's **Local tools** section lists anything missing. After you install a tool or change a path, click **Check again**.

## Generating from the overlay

1. Open a local video in mpv and select its Japanese audio track.
2. Press `Ctrl+Shift+G`. If the subtitle sidebar is empty, its **Generate Japanese subtitles** button opens the same modal.
3. Pick a model and download it if needed.
4. Optionally check **Focus on spoken dialogue** (see below).
5. Click **Generate subtitles**.

The modal shows progress. **Cancel** stops the job. Closing the modal lets the job keep running, and reopening it shows the progress.

SubMiner saves `<video>.ja.generated.srt` next to the video and adds a number if that name is taken. If the same file is still playing, it loads the subtitles and resets the subtitle delay.

Change the shortcut with `shortcuts.openSubtitleGeneration`.

## Generating from the launcher

```bash
subminer generate-subs                                  # current mpv file and audio track
subminer generate-subs episode.mkv --download-model
subminer generate-subs episode.mkv --model-path /path/to/ggml-small.bin
```

| Flag                     | What it does                                                  |
| ------------------------ | ------------------------------------------------------------- |
| `--model <name>`         | Use this managed model, such as `small` or `large-v3-turbo`   |
| `--download-model`       | Download the managed model if it is missing                   |
| `--model-path <path>`    | Use an existing model file                                    |
| `--audio-stream <index>` | Pick an audio stream by its absolute FFmpeg index             |
| `--output <path>`        | Write to this SRT path. Existing files are never overwritten. |

With a file argument, SubMiner uses the audio stream tagged Japanese, or the first stream. `Ctrl+C` cancels.

## Choosing a model

The modal recommends **large-v3-turbo** if it finds an NVIDIA GPU (`nvidia-smi`) and your `whisper-cli` can use CUDA. Otherwise it recommends **small**. AMD, Vulkan, and Apple GPUs do not trigger the turbo recommendation. The recommendation does not change your settings.

| Model                  | Tradeoff                                        |
| ---------------------- | ----------------------------------------------- |
| tiny, base             | Fast and small, more recognition errors         |
| small                  | Balanced quality and CPU time                   |
| medium, large-v1/v2/v3 | More accurate, needs more memory and time       |
| large-v3-turbo         | Faster than large-v3 with a small accuracy loss |

Quantized variants (`-q5_0`, `-q5_1`, `-q8_0`) use less disk and memory, with some accuracy loss. Your pick in the modal lasts for the session. Set `subtitleGeneration.managedModel` to change the default.

## Prioritizing spoken dialogue

**Focus on spoken dialogue** uses a speech detection (VAD) model to drop silent stretches before transcription. Long passages are split near detected speech, which reduces subtitles that appear before the line is spoken.

It needs two extra pieces:

- The Silero VAD model. Click **Download speech detection model** in the modal, or set `vadModelPath` to your own [Silero GGML model](https://huggingface.co/ggml-org/whisper-vad/tree/main).
- whisper.cpp's [speech segment detector](https://github.com/ggml-org/whisper.cpp/tree/master/examples/vad-speech-segments), found as `whisper-vad-speech-segments` or `vad-speech-segments` on `PATH`. Set `vadPath` for any other location.

The checkbox lasts for the session. Setting `vadModelPath` turns it on by default.

Dialogue mode keeps music and background sound that might contain speech, so songs can still produce subtitles. It can also take longer than a plain run, because each passage is transcribed separately.

## Using loaded subtitles as timing references

If the video playing in mpv already has a dialogue subtitle track loaded, SubMiner uses its cue times to decide where to split long audio. This works with or without dialogue mode. Whisper still writes the Japanese text and final timestamps. The launcher uses a reference only when its input is the file open in mpv.

SubMiner prefers English tracks and tracks labeled full or dialogue. It skips forced, image-based, generated, and signs or songs tracks, based on their titles and file names. An unlabeled signs-only file can slip through.

The reference must be timed correctly for the video. SubMiner does not fix a mistimed reference.

## Limitations

- Only local files and their internal audio tracks are supported. Not URLs, and not a separate audio file loaded in mpv. Pass a separate local audio file to the launcher instead.
- Whisper can miss, repeat, or invent lines, and its timing is approximate, especially over music or overlapping speech. Check the text and audio when you mine.
- SubMiner does not translate subtitles.
