# Bun 1.3.5 source and relinking materials

SubMiner redistributes unmodified Bun 1.3.5 executables from the official Bun release. The executables report revision `1e86cebd74a5723e818b5c0555276b646bcf0e4c` through `bun --revision`. Bun's `bun-v1.3.5` Git tag points to the next commit, `fa5a5bbe556a4bda5bde77b4013aa6c3bb4ec9ab`, so the tag archive does not exactly match the distributed executables.

The SubMiner GitHub release containing this application also contains `bun-v1.3.5-source.tar.gz` and its `.sha256` file. The archive contains:

- Bun source at the binary's reported revision
- WebKit and JavaScriptCore source at `6d0f3aac0b817cc01a846b3754b21271adedac12`
- TinyCC source at `29985a3b59898861442fa3b43f663fc1af2591d7`
- every external source repository registered by that Bun revision's CMake build, at its exact commit
- Bun's dependency patches, build scripts, lockfiles, collected third-party license files, and instructions for rebuilding Bun against a modified JavaScriptCore

The machine-readable inventory lives at `SOURCE-INVENTORY.json` inside the source archive and at `build/bun-source-manifest.json` in SubMiner's source repository.

The archive vendors the source repositories linked through Bun's CMake build. It preserves Bun's package-manager lockfiles and lol-html's `Cargo.lock`, but it does not vendor npm packages, crates.io packages, compilers, SDKs, or other build tools. Rebuilding needs network access for those package-manager and toolchain inputs. The archive's rebuild README records the known tool versions and the remaining unpinned Rust nightly input.

Bun's own license overview is included as `Bun-LICENSE.md`. WebKit's JavaScriptCore copy of GNU Library General Public License version 2 is included as `LGPL-2.0.txt`. TinyCC's GNU Lesser General Public License version 2.1 is included as `LGPL-2.1.txt`. `THIRD-PARTY-NOTICES.md` collects license texts from Bun's other externally fetched linked dependencies and identifies the scope of the remaining per-file notices in the source archive.

This notice describes the materials supplied with the release. It is not a legal-compliance or reproducible-build claim.
