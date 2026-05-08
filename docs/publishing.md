# Publishing

The repository has a separate **Release** GitHub Actions workflow (`.github/workflows/release.yml`) for publish readiness checks and real publishes.

## Dry run

Run **Release** manually (`workflow_dispatch`) with:

- `version=0.1.0-alpha.1`
- `dry-run=true`
- `publish-nuget=true` (ignored while `dry-run=true`)
- `publish-npm=true` (ignored while `dry-run=true`)
- `deploy-vercel=true` (ignored while `dry-run=true`)

Dry run validates the publish path without external credentials. Branch pushes to `clawy/**` also run the Release workflow in forced dry-run mode so PR changes can prove the workflow live before merge:

- Builds, tests, and packs `DocxSax` and `DocxSax.Tool` NuGet packages.
- Pushes the `.nupkg` files into a temporary local NuGet feed to exercise the push command path without hitting nuget.org.
- Builds/tests the single `docx-sax` npm package, including `/node` and `/browser` exports.
- Runs `npm publish --dry-run` and `npm pack` for `docx-sax`, then uploads the package artifact.
- Builds the Next.js WASM demo and checks that the Vercel CLI is available.

## Real publish inputs

Real publishing is deliberately manual and only happens from the Release workflow. Run Release with the exact `version`, `dry-run=false`, then opt in/out of each target:

- `publish-nuget=true` publishes NuGet packages.
- `publish-npm=true` publishes npm packages.
- `deploy-vercel=true` deploys the demo to Vercel production.

Required GitHub secrets:

- NuGet trusted publishing entry for `.github/workflows/release.yml` and packages `DocxSax`/`DocxSax.Tool`; set `NUGET_USER` to the nuget.org profile name used by that trusted-publishing policy (no long-lived `NUGET_API_KEY` required).
- npm trusted publishing entry for `.github/workflows/release.yml` and package `docx-sax` (no `NPM_TOKEN` or `NODE_AUTH_TOKEN` required for npm publish).
- `VERCEL_TOKEN` — Vercel token.
- `VERCEL_ORG_ID` — Vercel team/user id.
- `VERCEL_PROJECT_ID` — Vercel project id for `demos/nextjs-wasm`.

## Current package shape

- NuGet packages are the canonical first publish target.
- `docx-sax/browser` packages the browser WASM bridge and published runtime output.
- `docx-sax/node` currently ships a Linux x64 native bridge only; it is suitable for an alpha publish but should be treated as platform-limited until per-platform native assets/install strategy are added.
- The Vercel demo is a showcase/developer proof, not a required package dependency.
