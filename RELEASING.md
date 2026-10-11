# Releasing

A portable build publishes itself: push a tag and CI builds it, signs it, and
publishes the archive plus the `latest.json` that an installed copy checks.

## Cutting a version

Four things move together. Miss one and the build either fails outright or
publishes something it should not.

**1. The version, in three files.** `tauri build` refuses to run unless they
agree, so a mismatch fails the build rather than shipping something that
disagrees with itself.

| File | Field |
|---|---|
| `package.json` | `version` |
| `src-tauri/Cargo.toml` | `[package] version` |
| `src-tauri/tauri.conf.json` | `version` |

`Cargo.lock` records the package's own version as well, so it goes stale the
moment `Cargo.toml` changes. Any cargo command rewrites it — run one before
committing instead of editing it by hand.

**2. A section in each of the three changelogs** — `CHANGELOG.md`,
`CHANGELOG.zh-CN.md`, `CHANGELOG.ja.md`. The build refuses to run if any of them
is missing it, which is deliberate: a release with nothing to say about itself
is worse than no release, and a translation that quietly lags is worse than none
at all.

Those sections become the release page (all three, under a heading each) and the
notes the app shows when it offers the update. The app reads the notes for its
own locale, from the `notesLocales` map the build writes into `latest.json` —
`notes` stays English beside it, for any other updater client and as the
fallback for a release that predates the map.

Write for whoever is being asked to install it, not for whoever wrote the code.

**3. Commit, tag, push.**

```bash
git add package.json src-tauri/Cargo.toml src-tauri/Cargo.lock CHANGELOG*.md
git commit -m "chore: release 0.1.2"
git tag v0.1.2
git push origin main && git push origin v0.1.2
```

**The tag must match `tauri.conf.json`.** CI takes the version from that file,
not from the tag. A `v0.1.3` tag on a build that still says `0.1.2` would go on
to re-upload over the *existing* `v0.1.2` release and overwrite its artifacts.

**CI builds the workflow as it stands at that tag, not as it stands on `main`.**
So a change to the release workflow only takes effect on the *next* version. The
v0.1.1 tag was cut before the changelog and the translated READMEs were wired
into the archive, which is how this trap was found — check the tag, not `main`,
when a release is missing something the workflow should have added.

## What the archive carries

`SheetLens.exe`, the `portable.sheetlens` marker, `update-manifest.json` — and
**all three READMEs and all three changelogs**, so the folder explains itself
without a network. If a release is missing them, the workflow it was built from
predates them.


## What a release needs that is not in the repo

The private half of the signing key, as the repository secret
`TAURI_SIGNING_PRIVATE_KEY`. Without it the signing step fails and nothing is
published. The keypair also sits in `signing/`, which is ignored — that copy is
the backup, not what CI reads. **Losing it ends updates for every copy already
installed.**

The workflow needs `contents: write`, and the repository's Actions default has
to allow it. That default was `read`, which made every release fail at the last
step even though the workflow asks for write in its own `permissions:` block —
set to `write` on 2026-10-11. It is a repository setting, not a file, so nothing
in a checkout records it:

```bash
gh api repos/wangjianguo0405/sheet-lens/actions/permissions/workflow \
  -X PUT -f default_workflow_permissions=write
```

That default now grants write to **any** workflow that does not declare
`permissions:` of its own, so a new workflow should say what it actually needs
rather than inheriting this.

## When a release fails

Each of these has happened once, and the message does not name its cause:

| What you see | What it is |
|---|---|
| `Repository secret TAURI_SIGNING_PRIVATE_KEY is not set.` | the signing secret is missing or was rotated |
| `HTTP 403: Resource not accessible by integration` on the create, then `release not found` | the token came back read-only — see the repository setting above |
| `CHANGELOG.<xx>.md has no '## [x.y.z]' section` | a changelog is missing the version being released |
| the artifacts under an older release changed | the tag and `tauri.conf.json` disagree — see above |


## What a version bump does not touch

- `.github/workflows/release.yml` — it reads the version from `tauri.conf.json`
  and derives the tag, the asset name and `latest.json` from it.
- The updater config, the public key, the marker file. All version-independent.
- `latest.json`, the `.sig`, the zip — generated per release.

## After the release

An installed portable copy checks **once, at startup** — nothing polls. So a
copy that is already running will not notice a new version until it is
restarted, or its update dialog is opened (opening it re-checks).

Applying an update needs two releases: install one, publish the next, then
update. Note that the update UI belongs to the *running* build, not the
downloaded one — a build released before a UI change will not show it while
updating away from itself.
