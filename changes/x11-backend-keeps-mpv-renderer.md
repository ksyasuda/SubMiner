type: fixed
area: playback

- XWayland/X11 mode (`--backend=x11`, or the automatic fallback on non-Hyprland/Sway Wayland sessions) no longer forces mpv onto `--vo=gpu --gpu-api=opengl`. It now only pins the window context (`--gpu-context=x11vk,x11egl,x11`), so a `vo=gpu-next` config keeps its renderer, API, and user shaders. Forcing the legacy OpenGL renderer crashed mpv on the first fullscreen toggle for anyone using a gpu-next user shader that emits a 4-component LUMA hook (ArtCNN and friends), which asserts in mpv's old renderer (`copy_image: *offset + count < sizeof(dst)`) as soon as the shader's upscale-only condition turns on.
