type: added
area: subtitles

- Generate local Japanese SRT subtitles with whisper.cpp from a standalone modal opened with Ctrl+Shift+G, the subtitle sidebar button, or `subminer generate-subs`, with shared progress reporting, cancellation, safe output files, and automatic loading into the matching mpv video.
- Configure an existing multilingual model in Settings or choose an official multilingual model, including quantized variants, in the modal or launcher. The modal shows download sizes, speed and accuracy guidance, and a recommended starting model before explicitly downloading a verified SubMiner-managed model. Executable paths are optional overrides; empty fields find whisper-cli, ffmpeg, and ffprobe on PATH.
- Optionally select Prioritize dialogue in the modal and download the separate Silero speech detection model with progress and cancellation. The choice lasts for the session; a configured VAD model path sets the default. With the detector executable installed, transcribe separate speech passages at their original positions, reset text context between passages, and keep subtitle cues within those passages instead of spanning music breaks.
