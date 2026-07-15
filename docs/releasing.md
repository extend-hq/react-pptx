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

When the `NPM_PUBLISH_ENABLED` repository Actions variable is `true`, a successful `CI` run
for `main` triggers `.github/workflows/publish.yml`. The release workflow checks whether that
exact version already exists on npm. If it is new, an unprivileged job builds and verifies one
tarball, then a minimal privileged job publishes those exact bytes, pushes a matching
`v<version>` Git tag, and creates a GitHub release. A prerelease such as `0.2.0-beta.0` is
published under the matching `beta` npm dist-tag.

Do not create the Git tag manually. A tag is created only after npm publishing succeeds.
Keep `NPM_PUBLISH_ENABLED` unset or `false` until npm ownership and authentication are ready.

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

npm requires the package to exist before a Trusted Publisher can be attached. For the
first release only:

1. Add a GitHub Actions secret named `NPM_TOKEN` containing a granular npm token allowed
   to publish `@extend-ai/react-pptx`.
2. Add a repository Actions variable named `NPM_PUBLISH_USE_TOKEN` with value `true`.
3. Add a repository Actions variable named `NPM_PUBLISH_ENABLED` with value `true`.
4. Push or merge the release commit to `main`. If that commit already passed before publishing
   was enabled, manually run the `CI` workflow on `main`; its successful completion triggers
   `Publish Packages` for the same commit.
5. Configure the npm Trusted Publisher using the values above.
6. Delete `NPM_TOKEN`, delete or set `NPM_PUBLISH_USE_TOKEN` to `false`, and revoke the
   bootstrap token on npm.

The package is public, so successful OIDC releases also receive npm provenance.
