# Milton Releases Process Design

## Overview

This document describes the proposed release process for the Milton Office add-in.

The goal is to serve the Excel task pane, commands page, and static resources from a stable HTTPS origin while keeping builds reproducible, rollbacks easy, and Office client caching behavior predictable.

The recommended hosting model is:

```text
GitHub
  source, tags, build workflow, release artifacts, provenance

Cloudflare Pages
  static hosting, cache headers, production deployment, emergency rollback
```

GitHub remains the source of truth. Cloudflare is the static delivery layer.

---

## Goals

- Keep Office manifest URLs stable across normal application releases.
- Avoid changing `SourceLocation` URLs for routine task pane updates.
- Serve task pane and commands HTML with revalidation-friendly cache headers.
- Serve content-hashed JavaScript and CSS bundles as immutable assets.
- Build from a pinned source revision with a locked dependency graph.
- Preserve exact release artifacts for later redeployments and rollbacks.
- Support fast emergency rollback through Cloudflare Pages.
- Support auditable rollback by redeploying a previously published GitHub Release artifact.
- Avoid embedding private model provider secrets in browser bundles.

## Non-Goals

- Build-time static site generation or hydration in the initial release process.
- Runtime server-side rendering.
- Hosting application secrets in the task pane.
- Rebuilding an old tag as the primary rollback mechanism.
- Changing Office add-in manifests for every release.

---

## Core Release Model

Office should point at stable production URLs:

```text
https://addin.example.com/taskpanes/excel/taskpane.html
https://addin.example.com/commands/commands.html
```

The manifest should not point at versioned task pane URLs such as:

```text
https://addin.example.com/releases/v0.3.1/taskpanes/excel/taskpane.html
```

Versioned manifest URLs make Office client caching and tenant deployment harder to reason about. The stable HTML page should be the update control point.

Each release still produces immutable release artifacts:

```text
milton-office-v0.3.1-dist.tar.gz
excel.release.xml
sha256sums.txt
release.json
optional provenance / attestation files
```

Production is a promotion of one built artifact to the stable Cloudflare Pages project.

---

## URL And Cache Policy

The release site should use a custom domain rooted at the add-in origin:

```text
https://addin.example.com/
```

Using a root origin keeps Vite `base` simple:

```text
base: "/"
```

The cache policy should follow the normal web application pattern:

```text
/taskpanes/excel/taskpane.html      stable URL, revalidate before reuse
/commands/commands.html             stable URL, revalidate before reuse
/release.json                       stable URL, revalidate before reuse
/assets/*.js                        content-hashed, immutable
/assets/*.css                       content-hashed, immutable
```

Do not apply long-lived immutable caching to every file under `/assets/*` unless every asset in that directory is content-hashed. The current project copies static public assets such as icons and logo files with stable names. Those should either:

- keep a shorter cache policy, or
- move to content-hashed imported assets before receiving immutable cache headers.

Example Cloudflare Pages `_headers` file:

```text
/taskpanes/excel/taskpane.html
  Cache-Control: no-cache, must-revalidate

/commands/commands.html
  Cache-Control: no-cache, must-revalidate

/release.json
  Cache-Control: no-cache, must-revalidate

/assets/*.js
  Cache-Control: public, max-age=31536000, immutable

/assets/*.css
  Cache-Control: public, max-age=31536000, immutable

/assets/*
  Cache-Control: public, max-age=600
```

This lets the stable HTML document discover the current content-hashed bundle names while allowing repeat visits to reuse bundles aggressively.

---

## Manifest Strategy

The repo should keep separate local and release manifest paths:

```text
apps/office/manifests/excel.local.xml
apps/office/manifests/excel.release.xml
```

The local manifest points at:

```text
https://localhost:3000/
```

The release manifest points at:

```text
https://addin.example.com/
```

Routine application releases should not require manifest changes. Manifest updates should be reserved for changes to:

- add-in ID or version metadata
- permissions
- ribbon commands
- icon URLs
- support URLs
- app domains
- source URL shape

The release manifest should be validated in CI and attached to every GitHub Release. That makes the deployed add-in state auditable even when the manifest is distributed through sideloading or centralized deployment outside of GitHub.

---

## Build Reproducibility

Release builds should run in GitHub Actions from a tag or selected commit.

The build should pin:

- Node version
- pnpm version through `packageManager`
- dependency versions through `pnpm-lock.yaml`
- Vite production mode
- Cloudflare Pages project name through GitHub variables

The build should use:

```sh
corepack enable
pnpm install --frozen-lockfile
pnpm --filter @milton/office build
```

The workflow should avoid embedding volatile values into the JavaScript bundle. If build metadata is needed, write it to a separate `release.json` file after the app build. If byte-for-byte reproducibility becomes a strict requirement, keep `release.json` deterministic too, or exclude it from the deterministic app artifact hash.

Browser-exposed Vite environment variables are public. Do not inject private provider keys such as `DEBUG_OPENAI_API_KEY` into release builds.

---

## Release Artifacts

Each release should publish a GitHub Release containing:

```text
milton-office-vX.Y.Z-dist.tar.gz
excel.release.xml
sha256sums.txt
release.json
```

Recommended `release.json` shape:

```json
{
  "name": "milton-office",
  "version": "v0.3.1",
  "gitSha": "0123456789abcdef",
  "channel": "production"
}
```

The deployed site should also expose `/release.json`, so production can be inspected without guessing which artifact is live.

The release tarball is the rollback unit. Rollbacks should redeploy the old tarball, not rebuild old source by default.

---

## GitHub Actions And Secrets

Cloudflare deployment credentials should be stored in GitHub Actions secrets or environment secrets.

Recommended GitHub Environment:

```text
production
```

Recommended secrets:

```text
CLOUDFLARE_API_TOKEN
```

Recommended variables:

```text
CLOUDFLARE_ACCOUNT_ID
CLOUDFLARE_PAGES_PROJECT
PUBLIC_BASE_URL
```

`CLOUDFLARE_ACCOUNT_ID` is not intrinsically secret, but storing it as a secret is also acceptable if the team prefers keeping all deployment identifiers together.

The Cloudflare API token should be scoped as narrowly as possible. For Cloudflare Pages deploys, the intended permission is account-level Cloudflare Pages edit access.

The production GitHub Environment should require approval before deploy jobs can access production deployment credentials.

---

## Deployment Workflow

The release workflow should separate building from deployment:

```text
build-release
  checkout tag
  install dependencies with frozen lockfile
  validate local and release manifests
  build Office app
  add _headers
  add release.json
  create dist tarball
  create sha256sums.txt
  upload GitHub Release assets

deploy-production
  requires production environment approval
  download the exact release tarball
  unpack it
  deploy unpacked dist to Cloudflare Pages
```

The deploy step should use Cloudflare Pages Direct Upload through Wrangler:

```sh
wrangler pages deploy apps/office/dist --project-name "$CLOUDFLARE_PAGES_PROJECT"
```

Using Direct Upload keeps the build in GitHub Actions, where it is easier to pin tool versions, attach artifacts, and audit source-to-artifact provenance.

---

## Rollback Strategy

There are two rollback paths.

### Fast Rollback

Use Cloudflare Pages native rollback to restore the previous production deployment.

This is the emergency path when production is broken and time matters.

### Auditable Rollback

Run a GitHub Actions rollback workflow with an input release tag:

```text
release_tag = v0.3.0
```

The workflow should:

```text
download milton-office-v0.3.0-dist.tar.gz from the GitHub Release
verify sha256sums.txt
deploy that exact artifact to Cloudflare Pages
record the deployment in GitHub
```

This avoids drift from toolchain, registry, lockfile, or environment changes.

Rollback does not require changing the Office manifest because the source URLs remain stable.

---

## Preview And Staging

The simplest initial setup is:

```text
main branch
  build and validate only

pull requests
  build and validate only

tags
  build release artifact
  deploy to production after environment approval
```

If preview deployments are useful, use either:

- a separate Cloudflare Pages project for previews, or
- Cloudflare Pages preview deployments from non-production branches.

Preview URLs should not be placed in production manifests. They are for development and QA only.

---

## Security Notes

- Do not put private API keys into Vite release environment variables.
- Do not add `X-Frame-Options: DENY`; Office needs to host the page in its task pane WebView.
- Be careful with `Content-Security-Policy` `frame-ancestors`; Office embedding must remain allowed.
- Keep Cloudflare deploy tokens scoped to Pages deployment only.
- Prefer GitHub Environment approval for production deployments.
- Pin GitHub Actions by major version initially; consider pinning by SHA if supply-chain hardening becomes a priority.

---

## Implementation Plan

### Phase 1: Release Hosting Foundation

Create the Cloudflare Pages project as a Direct Upload project.

Configure a custom domain for the add-in origin:

```text
https://addin.example.com/
```

Add GitHub Environment configuration:

```text
environment: production
secret: CLOUDFLARE_API_TOKEN
variable: CLOUDFLARE_ACCOUNT_ID
variable: CLOUDFLARE_PAGES_PROJECT
variable: PUBLIC_BASE_URL
```

Add required reviewers for the production environment.

Outcome:

```text
Cloudflare can receive deploys from GitHub Actions, but no app workflow exists yet.
```

### Phase 2: Release Manifest And Build Configuration

Add a release manifest:

```text
apps/office/manifests/excel.release.xml
```

Point it at the production origin:

```text
https://addin.example.com/taskpanes/excel/taskpane.html
https://addin.example.com/commands/commands.html
https://addin.example.com/assets/icon-32.png
```

Update the Vite config so release builds can set the public base path, while local dev remains unchanged.

Add package scripts such as:

```json
{
  "build:release": "pnpm --filter @milton/office build",
  "validate:release": "office-addin-manifest validate apps/office/manifests/excel.release.xml"
}
```

Outcome:

```text
The repo can build a release artifact whose URLs match the Cloudflare production origin.
```

### Phase 3: Cache Headers And Release Metadata

Add Cloudflare Pages headers to the built output.

Preferred source path:

```text
apps/office/public/_headers
```

Add cache rules for:

```text
HTML: no-cache, must-revalidate
release.json: no-cache, must-revalidate
hashed JS/CSS: public, max-age=31536000, immutable
stable-name assets: short cache
```

Add a small release metadata generation step that writes:

```text
apps/office/dist/release.json
```

Outcome:

```text
Production can be inspected at /release.json, and browser caching follows the intended stable-HTML plus immutable-bundles model.
```

### Phase 4: Build And Artifact Workflow

Add a GitHub Actions workflow for tags:

```text
.github/workflows/release-office.yml
```

The workflow should:

```text
checkout source
enable corepack
install with pnpm --frozen-lockfile
validate manifests
build the Office app
write release.json
create dist tarball
create sha256sums.txt
attach artifacts to the GitHub Release
```

Do not deploy in the first version of this workflow until artifact generation is verified.

Outcome:

```text
Every release tag produces auditable release artifacts.
```

### Phase 5: Production Deployment Workflow

Extend `release-office.yml` or add:

```text
.github/workflows/deploy-office.yml
```

The deploy job should:

```text
wait for production environment approval
download the just-created release artifact
verify checksums
deploy dist to Cloudflare Pages with wrangler
publish the Cloudflare deployment URL in the GitHub deployment record
```

Outcome:

```text
Approved release tags deploy to the stable Office add-in origin.
```

### Phase 6: Rollback Workflow

Add:

```text
.github/workflows/rollback-office.yml
```

Inputs:

```text
release_tag
```

The workflow should:

```text
download the selected GitHub Release artifact
verify checksums
require production environment approval
deploy the exact artifact to Cloudflare Pages
record the rollback deployment
```

Outcome:

```text
Rollback is an artifact promotion operation, not a rebuild.
```

### Phase 7: Verification And Operational Runbook

Add release checks:

```text
curl production /release.json
curl taskpane.html and verify Cache-Control
curl generated JS/CSS and verify Cache-Control
validate release manifest
smoke load the task pane URL outside Office
```

Document runbook commands for:

```text
cutting a release
approving a deploy
rolling back through Cloudflare
rolling back through GitHub artifact redeploy
checking production version
```

Outcome:

```text
The release process is repeatable and operable without reconstructing tribal knowledge.
```

---

## Future Improvements

### Build-Time SSG And Hydration

Later, Milton can prerender the task pane shell at build time and hydrate it in the browser.

This would improve time to first paint by serving useful HTML before the React bundle finishes loading. The initial release process should defer this because it requires careful handling of hydration-safe state and Office-specific browser APIs.

When revisited, the likely direction is:

```text
build-time render task pane shell to HTML
hydrate with React hydrateRoot
guard Office APIs behind client-only code
keep generated HTML cache policy as no-cache
keep generated JS/CSS bundles immutable
```

### Code Splitting For Startup

The current task pane should avoid loading agent/model-provider code before the user can see the app shell. A later performance pass should split interactive agent code away from initial app chrome.

### Stronger Provenance

Add GitHub artifact attestations and optionally an SBOM for each release artifact.

### Canary Channel

Add a second stable origin:

```text
https://canary-addin.example.com/
```

This would allow a canary manifest to track upcoming releases without affecting production users.

### Toolchain Container

If stronger build reproducibility is needed, run release builds in a pinned container image with a fixed Node and pnpm toolchain.

### Signed Manifest Distribution

If centralized deployment becomes more formal, document how release manifests are reviewed, versioned, and distributed to Office tenants.

---

## References

- Cloudflare Pages custom headers: https://developers.cloudflare.com/pages/configuration/headers/
- Cloudflare Pages Direct Upload with CI: https://developers.cloudflare.com/pages/how-to/use-direct-upload-with-continuous-integration/
- GitHub Actions secrets: https://docs.github.com/en/actions/concepts/security/secrets
- Office Add-in SourceLocation: https://learn.microsoft.com/en-us/javascript/api/manifest/sourcelocation
- Office Add-in cache clearing: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache
- Vite static deploy guidance: https://vite.dev/guide/static-deploy
