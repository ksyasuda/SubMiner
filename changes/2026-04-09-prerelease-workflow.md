type: internal
area: release

- Added a dedicated beta/rc prerelease GitHub Actions workflow that publishes GitHub prereleases without consuming pending changelog fragments or updating AUR.
- Added prerelease note generation so beta and release-candidate tags can reuse the current pending `changes/*.md` fragments while leaving stable changelog publication for the final release cut.
