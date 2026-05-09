# Releasing

This project ships releases automatically. As a maintainer you only ever
need to merge pull requests; the rest is handled by GitHub Actions.

## How a release happens

1. **You merge feature PRs to `main`** with [Conventional Commit](https://www.conventionalcommits.org/)
   titles (`feat:`, `fix:`, `feat!:` for breaking changes, etc.).
2. **`release-please` opens a Release PR** that bumps the version in
   `visiowings/__init__.py`, `setup.py`, and `.release-please-manifest.json`,
   and updates `CHANGELOG.md` from the new commits.
3. **You merge the Release PR.** That creates the `vX.Y.Z` tag, publishes
   a GitHub Release, and triggers `.github/workflows/publish.yml`.
4. **`publish.yml` runs**:
   - Builds wheel + sdist (`python -m build`), validates with `twine check`.
   - Builds the standalone Windows EXE via PyInstaller, smoke-tests it,
     and attaches it to the Release.
   - Generates a CycloneDX SBOM (`sbom.cdx.json`) and a license report
     (`licenses.json`).
   - **Signs both with Sigstore** via keyless OIDC, producing
     `sbom.cdx.json.sigstore` and `licenses.json.sigstore` bundles, and
     attaches them to the Release.
   - Uploads the wheel + sdist to PyPI via OIDC Trusted Publishing — no
     long-lived API token required.

The whole pipeline is idempotent and re-runnable through
`workflow_dispatch`.

## Verifying a release

Every release artefact can be verified offline by anyone — no maintainer
key required, since Sigstore uses short-lived certificates tied to the
GitHub Actions OIDC identity.

```bash
pip install sigstore     # one-time

# Download the artefact and its .sigstore bundle from the Release page,
# then verify:
sigstore verify github \
    --bundle sbom.cdx.json.sigstore \
    --cert-identity 'https://github.com/twobeass/visiowings/.github/workflows/publish.yml@refs/tags/vX.Y.Z' \
    --cert-oidc-issuer 'https://token.actions.githubusercontent.com' \
    sbom.cdx.json
```

A successful run prints `OK: sbom.cdx.json` and proves the file was
produced by `publish.yml` running on the `vX.Y.Z` tag.

## Supply-chain posture: OpenSSF Scorecard

`.github/workflows/scorecard.yml` runs the
[OpenSSF Scorecard](https://github.com/ossf/scorecard) every Monday and
on every push to `main`. Results are uploaded to GitHub
**code-scanning** (Security tab → Code scanning alerts) so regressions
in branch protection, pinned actions, signed releases, etc. show up
alongside CodeQL findings.

To run a local Scorecard against the public repo:

```bash
docker run --rm -e GITHUB_AUTH_TOKEN=$GITHUB_TOKEN \
    gcr.io/openssf/scorecard:stable \
    --repo=github.com/twobeass/visiowings
```

## One-time setup: PyPI Trusted Publishing

This needs to be done **once** per project, before the first automated
release lands on PyPI. After that, GitHub Actions can publish without any
secret tokens.

1. Sign in at <https://pypi.org/> as an account that owns (or will own)
   the `visiowings` project. If the project does not exist yet, create
   the **pending publisher** first; PyPI will reserve the name for you.
2. Open the account settings: <https://pypi.org/manage/account/publishing/>
3. Click **Add a new pending publisher** with these values:

   | Field | Value |
   | --- | --- |
   | PyPI Project Name | `visiowings` |
   | Owner | `twobeass` |
   | Repository name | `visiowings` |
   | Workflow name | `publish.yml` |
   | Environment name | `pypi` |

4. Save. The first time `publish.yml` runs against this configuration,
   PyPI promotes the pending publisher to a configured trusted publisher
   and uploads the dists.

If you decide to publish to TestPyPI first, repeat the steps at
<https://test.pypi.org/manage/account/publishing/> and add a separate
job to `publish.yml` keyed off the `testpypi` environment.

## Cutting a manual release (escape hatch)

If `release-please` is broken or you need to ship out of band:

```bash
# 1. Bump the version in pyproject.toml, visiowings/__init__.py, setup.py
# 2. Update CHANGELOG.md
# 3. Tag and push
git tag vX.Y.Z
git push origin vX.Y.Z

# 4. Create a Release on GitHub for that tag
#    -> publish.yml will run automatically
```

## Yanking a bad release

If a release is broken on PyPI, yank it (it stays installable for users
who explicitly request that version, but disappears from the resolver):

```bash
# Web: https://pypi.org/manage/project/visiowings/release/X.Y.Z/
# CLI:
pip install twine
twine yank visiowings X.Y.Z --reason "regression in <area>"
```

Then ship a fix release immediately.
