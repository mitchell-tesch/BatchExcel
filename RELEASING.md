# Releasing BatchExcel

Releases are driven by **git tags**. Pushing a tag matching `v*` triggers
`.github/workflows/release.yml`, which builds, tests, publishes win-x64 binaries
(self-contained + framework-dependent single-file), zips them, computes
SHA-256 checksums, and creates a GitHub Release.

Tags containing a hyphen (semver pre-release, e.g. `v0.9.0-beta.1`) are
automatically marked as **pre-release** on GitHub.

## Beta release checklist

1. **Make sure `main` is green** (CI passes, `dotnet test` clean locally).
2. **Bump the version** in `BatchExcel/BatchExcel.csproj` so the in-repo
   value matches the tag you're about to push:
   - `<Version>` / `<InformationalVersion>` → full SemVer, e.g. `0.9.0-beta.1`
   - `<AssemblyVersion>` / `<FileVersion>` → 4-part numeric, e.g. `0.9.0.0`
   (CI also overrides these from the tag at publish time, but keeping the csproj
   in sync makes local builds report the right version.)
3. **Update `CHANGELOG.md`** (optional — the release uses
   `generate_release_notes: true` which auto-summarises commits / PRs since the
   previous tag).
4. **Commit & push**:
   ```powershell
   git add BatchExcel/BatchExcel.csproj CHANGELOG.md
   git commit -m "Release v0.9.0-beta.1"
   git push
   ```
5. **Tag & push the tag**:
   ```powershell
   git tag -a v0.9.0-beta.1 -m "BatchExcel 0.9.0-beta.1 (beta)"
   git push origin v0.9.0-beta.1
   ```
6. **Watch the workflow** at
   `https://github.com/mitchell-tesch/BatchExcel/actions` — when it finishes the
   release will appear under **Releases** marked **Pre-release**, with:
   - `BatchExcel-<version>-win-x64-self-contained.zip` — no .NET install required
   - `BatchExcel-<version>-win-x64-framework-dependent.zip` — needs .NET 10 Desktop Runtime
   - `SHA256SUMS.txt`
7. **Smoke-test** the self-contained zip on a clean Windows box (extract,
   launch `BatchExcel.exe`, run the bundled `Example/BatchExcel.xlsx`).
8. **Announce** with a link to the Release page. Ask testers to file issues
   tagged `beta` with the version string from **Help → About**.

## If a beta needs a fix

Bump to the next pre-release tag (`v0.9.0-beta.2`, etc.) and repeat. Never
move or force-push an existing release tag — published zips are immutable from
users' perspective.

## Promoting to a stable release

Tag without a pre-release suffix (`v1.0.0`). The workflow will publish it as
a normal (non-pre-release) GitHub Release.

## Undoing a bad release

```powershell
# Delete the GitHub Release + tag (also via the web UI)
gh release delete v0.9.0-beta.1 --yes
git push --delete origin v0.9.0-beta.1
git tag -d v0.9.0-beta.1
```

