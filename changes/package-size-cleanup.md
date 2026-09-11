type: changed
area: release

- Reduced installer and unpacked app size by excluding documentation demo media, dependency source maps, TypeScript files, test and fixture directories, other development files, and unused Koffi platform binaries, and sharing the existing Japanese UI font across windows.
- Added package content checks, published size reports with release comparisons, and packaged asset/native-module smoke checks to the shared stable and prerelease build workflow. Size growth is reported without blocking releases.
