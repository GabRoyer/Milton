# Upstream Diff

This package is imported from `earendil-works/pi`, upstream path `packages/agent`.

Current baseline:

- Repository: https://github.com/earendil-works/pi
- Tag: `v0.75.4`
- Commit: `3533843dd781dcd233f51854fc883ec246a6a919`
- Local path: `packages/pi-agent-core`

## Local Changes

### Source-consumption package manifest

`package.json` is adapted so Milton can consume this package directly from
TypeScript source inside the workspace:

- `main`, `types`, and package `exports` point at `src/*.ts` instead of `dist/*.js`.
- Added `@earendil-works/pi-agent-core/browser` for browser-safe imports.
- Added `@earendil-works/pi-agent-core/types` for type-only consumers.
- Changed `@earendil-works/pi-ai` to `workspace:*`.
- Changed `build` to `tsc --noEmit -p tsconfig.json`.
- Trimmed local dev dependencies to the TypeScript checker and Vitest runner
  needed for this workspace.

### Browser entrypoint

Added `src/browser.ts` as the browser-safe public entrypoint. It exports the low
level agent loop and related types without exporting Node-specific helpers from
`src/node.ts`, `src/proxy.ts`, or the durable harness.

### Browser-safe agent loop imports

`src/agent-loop.ts` imports only the needed `pi-ai` subpaths:

- `@earendil-works/pi-ai/types`
- `@earendil-works/pi-ai/utils/event-stream`
- `@earendil-works/pi-ai/utils/validation`

This avoids importing the full `@earendil-works/pi-ai` barrel into the browser
bundle.

### Required stream function injection

The local browser-compatible `agentLoop` no longer falls back to
`streamSimple` from `@earendil-works/pi-ai`. Callers must provide a `StreamFn`.
This keeps provider registration and Node-oriented code out of browser builds
and lets the app choose the exact browser-compatible provider stream.

`src/types.ts` defines `StreamFn` directly using `Model`, `Context`,
`SimpleStreamOptions`, and `AssistantMessageEventStream` from
`@earendil-works/pi-ai/types` instead of deriving it from `streamSimple`.

### Structured thrown tool error details

`src/agent-loop.ts` preserves a thrown tool error's `details` payload when
converting the thrown error into an `AgentToolResult`.

Upstream converted thrown tool failures into text content with empty
`details`. Milton needs structured details from the OfficeJS execution tool so
compile diagnostics, logs, elapsed time, and submitted source metadata can reach
`tool_execution_end` and taskpane debugging UI even when the tool fails.

The local behavior is limited to error paths in tool execution and
`afterToolCall`: when the thrown value exposes a non-`undefined` `details`
property, that value is copied into the generated error result; otherwise the
result keeps the upstream empty details object.

### Local TypeScript config

Added `tsconfig.json` for source type-checking in this workspace. It checks the
browser entrypoint with bundler-style module resolution.

### Changelog placeholder

`CHANGELOG.md` currently has an empty `[Unreleased]` section added locally.
