type: fixed
area: Anki media

- Fixed sentence-audio generation timing out on slow network-mounted MKV files with many subtitle and font-attachment streams. Selected audio tracks now use bounded FFmpeg probing and a two-minute extraction budget, and missing output reports a clear FFmpeg error instead of raw `ENOENT`.
