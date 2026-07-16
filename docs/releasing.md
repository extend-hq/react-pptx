# Releasing `@extend-ai/react-pptx`

The repository publishes one public npm package: `@extend-ai/react-pptx`. The model,
Wasm runtime, repository-owned renderer, worker, styles, and Wasm binary are bundled into that
package so consumers do not need internal workspace packages or repository patches.

## Normal release

1. Add a changeset for every user-visible package change:

   ```bash
   pnpm changeset
   ```

2. Apply the pending changesets and review the version and changelog:

   ```bash
   pnpm version-packages
   pnpm install --lockfile-only
   pnpm build
   pnpm pack:check
   ```

3. Commit and merge the version changes to `main`.

Changing `packages/react-viewer/package.json` on `main` directly triggers
`.github/workflows/publish.yml`. The release workflow checks whether that exact version already
exists on npm. If it is new, an unprivileged job builds, typechecks, and verifies one tarball,
stamps it with the release commit, then a minimal privileged job publishes those exact bytes,
pushes a matching
`v<version>` Git tag, and creates a GitHub release. A prerelease such as `0.2.0-beta.0` is
published under the matching `beta` npm dist-tag.

Do not create the Git tag manually. A tag is created only after npm publishing succeeds.
The workflow can also be started manually with `workflow_dispatch`; an already-published version
is verified and skipped rather than published again.

## npm Trusted Publishing

Normal releases use npm Trusted Publishing through GitHub Actions OIDC. Configure the
package on npm with:

- Provider: GitHub Actions
- Organization: `extend-hq`
- Repository: `react-pptx`
- Workflow filename: `publish.yml`
- Allowed action: `npm publish`
- Environment: none

The workflow filename is part of the npm trust relationship and must match exactly.

## First release bootstrap

npm requires the package to exist before a Trusted Publisher can be attached. The preferred
bootstrap is one publish from an authenticated npm CLI on the clean release commit:

1. Push the release commit and make sure `CI` passes for that exact commit.
2. Confirm the npm identity and build one verified tarball:

   ```bash
   npm whoami
   export REACT_PPTX_PACK_OUTPUT="$(mktemp -t react-pptx.XXXXXX.tgz)"
   pnpm build
   pnpm typecheck
   pnpm test
   pnpm pack:check
   ```

3. Publish the verified bytes and record npm's immutable hashes plus the release commit:

   ```bash
   npm publish "$REACT_PPTX_PACK_OUTPUT" --access public
   npm view @extend-ai/react-pptx@0.1.0 dist.shasum dist.integrity --json
   git rev-parse HEAD
   ```

4. Check `npm view @extend-ai/react-pptx@0.1.0 gitHead`. If npm reports the exact release
   commit, no bootstrap record is needed. If it is empty, add that package version, commit,
   `dist.shasum`, and `dist.integrity` to `.github/release-bootstrap.json`, then push the record
   through `CI`. This exception is hash-pinned and is only for a manual first publish.
5. Configure the npm Trusted Publisher using the values above.
6. Leave `NPM_PUBLISH_USE_TOKEN` unset or `false`.
7. Manually run `Publish Packages` on `main`. The publish workflow recognizes the existing npm version,
   verifies its `gitHead` or hash-pinned bootstrap record, creates the matching Git tag,
   and creates the GitHub release without republishing the package.

If an authenticated local publish is not available, the workflow also supports a one-time token
bootstrap:

1. Add a GitHub Actions secret named `NPM_TOKEN` containing a granular npm token allowed
   to publish `@extend-ai/react-pptx`.
2. Add a repository Actions variable named `NPM_PUBLISH_USE_TOKEN` with value `true`.
3. Push or merge the version-changing release commit to `main`. If that commit is already on
   `main`, manually run `Publish Packages` for the same commit.
4. Configure the npm Trusted Publisher using the values above.
5. Delete `NPM_TOKEN`, delete or set `NPM_PUBLISH_USE_TOKEN` to `false`, and revoke the
   bootstrap token on npm.

The package is public, so successful OIDC releases also receive npm provenance.
