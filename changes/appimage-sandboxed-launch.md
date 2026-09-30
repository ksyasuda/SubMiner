type: fixed
area: linux

- Fixed the AppImage exiting without opening a window when launched inside a sandbox that mounts the image itself, such as `firejail --appimage` (used by the AppImage catalog test). The detached background app now runs straight from the sandbox's squashfs mount instead of trying to re-mount the AppImage over FUSE.
