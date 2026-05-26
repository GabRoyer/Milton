# Upstream Diff

This package is imported from `earendil-works/pi`, upstream path `packages/ai`.

Current baseline:

- Repository: https://github.com/earendil-works/pi
- Tag: `v0.75.4`
- Commit: `3533843dd781dcd233f51854fc883ec246a6a919`
- Local path: `packages/pi-ai`

## Local Changes

### Source-consumption package manifest

`package.json` is adapted so Milton can consume this package directly from
TypeScript source inside the workspace:

- `main`, `types`, and package `exports` point at `src/*.ts` instead of `dist/*.js`.
- Added source subpath exports used by Milton packages:
  - `@earendil-works/pi-ai/models`
  - `@earendil-works/pi-ai/types`
  - `@earendil-works/pi-ai/utils/event-stream`
  - `@earendil-works/pi-ai/utils/validation`
- Removed the local `bin` entry that points at built `dist/cli.js`.
- Changed `build` to `tsc --noEmit -p tsconfig.json`.
- Trimmed local dev dependencies to the TypeScript checker needed for this workspace.
- Declared `@smithy/node-http-handler` directly because the upstream source imports it and pnpm source resolution should not rely on a transitive dependency.

### Local TypeScript config

Added `tsconfig.json` for source type-checking in this workspace. It uses
bundler-style module resolution and checks only the source entrypoints Milton
currently imports.

### Vite dynamic import guard

`src/env-api-keys.ts` keeps upstream's Node-only dynamic imports, but adds
`/* @vite-ignore */` so Vite does not try to statically analyze or bundle
`node:fs`, `node:os`, and `node:path` into the Office/browser bundle.
