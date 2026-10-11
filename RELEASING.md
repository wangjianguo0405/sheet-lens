# Releasing

A portable build publishes itself: push a tag and CI builds it, signs it, and
publishes the archive plus the `latest.json` that an installed copy checks.

## Cutting a version

The version lives in three files, and `tauri build` refuses to run unless they
agree — a mismatch fails the build rather than shipping something that
disagrees with itself. Bump all three to the same value:

| File | Field |
|---|---|
| `package.json` | `version` |
| `src-tauri/Cargo.toml` | `[package] version` |
| `src-tauri/tauri.conf.json` | `version` |

```bash
# after bumping all three
cd src-tauri && cargo check        # also refreshes Cargo.lock
git commit -am "chore: release 0.1.2"
git tag v0.1.2
git push origin main && git push origin v0.1.2
```

`Cargo.lock` records the package's own version as well, so it goes stale the
moment `Cargo.toml` changes. Any cargo command rewrites it — run one before
committing instead of editing it by hand.

**The tag must match `tauri.conf.json`.** CI takes the version from that file,
not from the tag. A `v0.1.3` tag on a build that still says `0.1.2` would go on
to re-upload over the *existing* `v0.1.2` release and overwrite its artifacts.

## What a release needs that is not in the repo

The private half of the signing key, as the repository secret
`TAURI_SIGNING_PRIVATE_KEY`. Without it the signing step fails and nothing is
published. The keypair also sits in `signing/`, which is ignored — that copy is
the backup, not what CI reads. **Losing it ends updates for every copy already
installed.**

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
