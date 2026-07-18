type: fixed
area: app

- Fixed "Service Crash" desktop notifications (KDE DrKonqi) after closing a video when running the Linux AppImage: the short-lived background bootstrap spawned a Chromium GPU child that outlived it (surviving `app.exit`) and died with SIGBUS at session end when the bootstrap's FUSE mount was finally released. The bootstrap now runs with the GPU in-process so it leaves no children behind, and the detached app's mount remains supervised until its Chromium children finish. Set `SUBMINER_NO_APPIMAGE_MOUNT_KEEPALIVE=1` to disable the detached-app mount supervisor.
