type: fixed
area: playback

- Fixed managed local playback so duplicate startup-ready retries no longer unpause media after a later manual pause on the same file.
- Fixed managed local subtitle auto-selection so local files reuse configured primary and secondary subtitle language priorities instead of staying on mpv's initial `sid=auto` guess.
