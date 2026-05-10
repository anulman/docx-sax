# Publishing

The repository has a separate **Release** GitHub Actions workflow (`.github/workflows/release.yml`) for publish readiness checks and real publishes.

## Dry run

Run **Release** manually (`workflow_dispatch`) with:

- `version=0.1.0-alpha.1`
- `dry-run=true`
- `publish-nuget=true` (ignored while `dry-run=true`)
- `publish-npm=true` (ignored while `dry-run=true`)

Dry run validates the publish path without external credentials:

- Builds, tests, and packs `DocxSax` and `DocxSax.Tool` NuGet packages.
- Pushes the `.nupkg` files into a temporary local NuGet feed to exercise the push command path without hitting nuget.org.
- Builds/tests/packs the npm workspace matrix: user-facing `@docx-sax/node` and `@docx-sax/browser`, implemented `@docx-sax/native-linux-x64`, and placeholder native payload packages `@docx-sax/native-darwin-arm64`, `@docx-sax/native-darwin-x64`, and `@docx-sax/native-win32-x64`.
- Runs `npm publish --dry-run` and `npm pack` for each npm package, then uploads one artifact per matrix row.
- Builds the Next.js WASM demo.

## Real publish inputs

Real publishing is deliberately manual and only happens from the Release workflow. npm publishes native payload packages before the user-facing adapters so dependency resolution is ready as soon as `@docx-sax/node` appears. Run Release with the exact `version`, `dry-run=false`, then opt in/out of each target.

The npm dist-tag is inferred from the version suffix: `1.2.3` publishes as `latest`, `1.2.3-alpha.1` as `alpha`, and `1.2.3-foo` as `foo`.

- `publish-nuget=true` publishes NuGet packages.
- `publish-npm=true` publishes npm packages.

Required GitHub secrets:

- NuGet trusted publishing entry for `.github/workflows/release.yml` and packages `DocxSax`/`DocxSax.Tool`; set `NUGET_USER` to the nuget.org profile name used by that trusted-publishing policy (no long-lived `NUGET_API_KEY` required).
- npm trusted publishing entries for `.github/workflows/release.yml` and packages `@docx-sax/node`, `@docx-sax/browser`, `@docx-sax/native-linux-x64`, `@docx-sax/native-darwin-arm64`, `@docx-sax/native-darwin-x64`, and `@docx-sax/native-win32-x64` (no `NPM_TOKEN` or `NODE_AUTH_TOKEN` required for npm publish).

## Current package shape

- NuGet packages are the canonical first publish target.
- `@docx-sax/node` and `@docx-sax/browser` are the v0 user-facing JavaScript packages.
- `@docx-sax/browser` packages the browser WASM bridge and published runtime output.
- `@docx-sax/native-linux-x64` currently ships the implemented Linux x64 native bridge.
- `@docx-sax/native-darwin-arm64`, `@docx-sax/native-darwin-x64`, and `@docx-sax/native-win32-x64` are honest placeholder prerelease packages for trusted-publisher/bootstrap only; they do not include runtime implementations yet.
- The Vercel demo is a showcase/developer proof, not a required package dependency.
