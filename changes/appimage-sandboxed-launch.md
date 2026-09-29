type: fixed
area: linux

- Fixed the AppImage exiting without opening a window when launched inside a sandbox that mounts the image itself, such as `firejail --appimage` (used by the AppImage catalog test). The detached background app no longer tries to re-mount the AppImage over FUSE when it is not running from a FUSE mount.
