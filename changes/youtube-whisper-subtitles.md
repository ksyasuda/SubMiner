type: added
area: youtube

- Added `youtube.subtitleSource`: set it to `whisper` to transcribe YouTube videos with Whisper instead of downloading YouTube's captions. The default stays `youtube`.
- Whisper mode keeps the video paused while it transcribes, with progress in the subtitle generation modal; close the modal to keep watching, or cancel to continue without subtitles.
- The subtitle generation modal (`Ctrl+Shift+G`) now also works on YouTube videos, whatever `youtube.subtitleSource` is set to.
- Whisper mode downloads a small audio-only stream and deletes it once generation ends, or when you switch videos, close mpv, or quit.
