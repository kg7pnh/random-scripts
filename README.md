# Random-Scripts

A collection of scripts written in different languages for various purposes.

## Scripts

| Script | What it does |
|---|---|
| [`powershell/Find-VideoFiles.ps1`](powershell/Find-VideoFiles.ps1) | Lists video files (or another kind, with `-Kind`) under a folder, using the Windows Search index. Example: `.\Find-VideoFiles.ps1 -Path 'D:\Videos'` |

## Releases

Releases are automated with [release-please](https://github.com/googleapis/release-please).
Write commit messages in the [Conventional Commits](https://www.conventionalcommits.org/) format:

- `fix: ...` makes a patch release
- `feat: ...` makes a minor release
- `feat!: ...` or a `BREAKING CHANGE:` footer makes a major release

After each push to `main`, release-please opens or updates a release PR.
Merging that PR updates `CHANGELOG.md`, tags the version (for example `v1.0.0`) and publishes a GitHub release.
