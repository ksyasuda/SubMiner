type: fixed
area: windows

- Acquire the app single-instance lock earlier so Windows overlay/video launches reuse the running background SubMiner process instead of booting a second full app and repeating startup warmups.
