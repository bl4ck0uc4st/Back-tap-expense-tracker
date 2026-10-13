# Contributing

Issues and pull requests are welcome.

## Before opening an issue

Check [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) first — it covers the
failures that come up most often, including several that produce no error
message at all.

If it is not covered, include:

- what you expected and what happened
- the output of `?echo=1` if the problem involves data not arriving
- the version string from any response
- iOS version, if the problem is on the Shortcuts side

## Pull requests

- Keep `src/Code.gs` as a single file. It is meant to be pasted into the Apps
  Script editor, so splitting it across modules makes setup harder.
- Bump the version string at the top when behaviour changes. Those markers are
  the fastest way to tell whether a deployment is stale.
- Comment anything non-obvious, particularly Sheets or Apps Script quirks. Most
  of this code exists because something failed silently.
- Do not commit a deployment URL. Anyone holding one can write to that sheet.

## Testing a change

1. Paste into Apps Script, save, deploy a new version.
2. Confirm the version marker in the response has changed.
3. Exercise a write via the browser before touching the Shortcut — it isolates
   script problems from Shortcut problems.
4. Run `?repair=1` and confirm every tab and chart rebuilds cleanly.
