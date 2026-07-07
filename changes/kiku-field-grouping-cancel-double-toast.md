type: fixed
area: anki

- Fixed cancelling the Kiku field grouping dialog showing two "Field grouping cancelled" notifications when grouping was started via the trigger shortcut: the manual workflow already notifies about its outcome (cancelled, UI unavailable, failed), and the trigger path re-notified on top of it. The workflow now owns all outcome notifications, and a previously silent failure (the original card no longer loadable) gets its own message.
