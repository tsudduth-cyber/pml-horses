# Issues handled

Kristina reports a problem in the PML Horse App project on Claude. When it
is the app's fault rather than a filing problem, that project writes a note
into `Horses/_app/Issues/` in SharePoint.

An hourly check reads that folder and works whatever is new. This file is
how it knows what it has already done — the SharePoint connector can read
that folder but not rename or move anything in it, so the record lives here
instead. One line per note, newest at the bottom.

Outcomes:

- `shipped` — fixed, tested and deployed.
- `stopped` — the fix touches something that writes back to SharePoint, so
  it is waiting on Taylor. That class destroyed Kristina's Still Missing
  notes on 2026-11-05 and does not go out unattended.
- `not a bug` — checked against the real files and the app is right.

| Note | Date | Outcome | What happened |
| --- | --- | --- | --- |
