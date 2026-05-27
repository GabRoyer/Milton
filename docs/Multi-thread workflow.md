# Multi-thread Workflow

## Overview

Milton development should support more than one active Codex thread at a time. Each thread should be able to work in its own Git worktree, run its own Office add-in build, and be debugged in Excel without clobbering another thread's add-in.

The core model is:

```text
Codex thread
  -> Git worktree
  -> derived dev identity
  -> generated Office manifest
  -> dedicated HTTPS localhost port
  -> separately sideloaded Excel add-in
```

The developer should not have to name profiles manually. The identity should be derived from the worktree path by default.

---

## Goals

- Let multiple Codex-created worktrees run Milton Office add-ins at the same time.
- Avoid manual profile naming for normal development.
- Make each worktree appear to Office as a distinct add-in.
- Keep add-in identity stable for the lifetime of a worktree.
- Keep generated development manifests out of source control.
- Preserve explicit overrides for unusual debugging cases.
- Make it obvious inside the task pane which worktree/build is currently running.

## Non-Goals

- Supporting parallel production manifests.
- Changing the release manifest strategy.
- Making Office run two versions of the same manifest identity at once.
- Eliminating Office cache clearing entirely. Manifest and ribbon changes can still require cache cleanup.

---

## Problem

The current local Office add-in uses a single manifest and a single Vite dev origin:

```text
apps/office/manifests/excel.local.xml
https://localhost:3000/
```

That is fine for one active debug session, but it does not scale to multiple worktrees. If two threads use the same manifest ID and the same `SourceLocation`, Office cannot reliably distinguish them. A task pane reload may also pick up the wrong Vite server if both worktrees expect port `3000`.

Parallel debugging needs isolation across three surfaces:

- Office manifest identity
- HTTPS localhost origin
- visible development label

---

## Core Design

Each worktree gets a generated local development identity. The checked-in manifest remains the source template, but the manifest used for sideloading is generated per worktree.

The default commands should stay simple:

```sh
pnpm dev
pnpm start
```

When run from different worktrees, these commands should resolve different development identities automatically.

The generated state should look conceptually like:

```text
apps/office/.generated/
  manifests/
    excel.local-8f3a21c4.xml
  dev-profile.json
```

These files are local build artifacts and should be ignored by Git.

---

## Path-derived Identity

The identity resolver should use the canonical Git worktree root as the default seed:

```sh
git rev-parse --show-toplevel
```

The implementation should canonicalize the path before hashing it. On macOS this should account for symlinks by resolving the real path. The same worktree path should always produce the same identity; a different worktree path should produce a different identity.

Recommended seed:

```text
realpath(git worktree root)
```

Recommended derived fields:

```text
profileKey      sha256(realpath)
shortProfile    first 8-12 hex chars of profileKey
displayName     Milton (<worktree basename>-<shortProfile>)
manifestId      UUIDv5(profileKey, Milton development namespace UUID)
defaultPort     deterministic port from profileKey
```

The generated manifest should replace at least:

- `<Id>`
- `<DisplayName DefaultValue="...">`
- `<IconUrl>`
- `<HighResolutionIconUrl>`
- `<SupportUrl>`
- `<AppDomain>`
- `<SourceLocation>`
- `Commands.Url`
- `Taskpane.Url`
- icon resource URLs

It is also reasonable to make command IDs and group IDs profile-specific if Excel proves sticky about cached command surfaces.

---

## Port Allocation

Ports should be automatic but stable.

Recommended behavior:

1. Compute a deterministic default port from the path-derived profile key.
2. Reuse the port already recorded for the current worktree profile, if one exists.
3. For a new profile, check whether the deterministic default port is already bound.
4. If the port is free, use it.
5. If the port is occupied or already registered by another active worktree, probe forward until a free port is found.
6. Persist the selected port for that worktree.

If `pnpm dev` is run while the selected port is already in use by the same generated profile, it should report that the dev server is already running. If the port is occupied by another process, it should fail with a clear message so the developer can stop that process, clean the profile, or set `MILTON_OFFICE_PORT`.

The port registry should be outside the worktree so multiple worktrees can coordinate:

```text
~/.config/milton/office-dev-profiles.json
```

Example registry shape:

```json
{
  "profiles": {
    "8f3a21c4": {
      "worktreeRoot": "/Users/groyer/Projects/Milton",
      "port": 3184,
      "manifestId": "..."
    }
  }
}
```

The checked-in default can remain `3000` for the primary worktree if desired, but generated worktree profiles should prefer a wider range such as `3100-3999` to avoid collisions with common local services.

Manual overrides should exist for emergencies:

```sh
MILTON_OFFICE_PORT=3177 pnpm dev
MILTON_DEV_PROFILE=demo-a pnpm start
```

Overrides should not be required for normal Codex thread work.

---

## Generated Manifest

The repository should keep a stable local manifest template. The sideloaded file should be generated.

Proposed files:

```text
apps/office/manifests/excel.local.xml
apps/office/scripts/dev-profile.mjs
apps/office/.generated/manifests/excel.<shortProfile>.xml
```

The generator should:

- derive the worktree identity
- allocate or reuse the port
- rewrite all localhost URLs
- rewrite the manifest ID
- rewrite the visible display name
- write the generated manifest
- write a small profile metadata file for scripts and UI

Example generated URLs:

```text
https://localhost:3184/taskpanes/excel/taskpane.html
https://localhost:3184/commands/commands.html
https://localhost:3184/assets/icon-32.png
```

The source template should stay human-readable. Avoid making developers edit generated XML directly.

---

## Script Behavior

The root scripts should continue to be the main entry points:

```sh
pnpm dev
pnpm start
pnpm stop
pnpm validate
```

Target behavior:

- `pnpm dev` resolves the current worktree profile and starts Vite on that profile's port.
- `pnpm start` resolves the current worktree profile, generates the manifest, and sideloads that generated manifest.
- `pnpm stop` resolves the current worktree profile and stops/unloads that generated manifest.
- `pnpm validate` validates the generated manifest for the current worktree.

The app package can expose lower-level commands for debugging:

```sh
pnpm --filter @milton/office manifest:generate
pnpm --filter @milton/office start
pnpm --filter @milton/office stop
pnpm --filter @milton/office validate
pnpm --filter @milton/office dev:profiles
pnpm --filter @milton/office dev:clean-profile
pnpm --filter @milton/office dev:clean-stale-profiles
```

The important behavior is that the default root commands remain path-aware and require no manual profile argument.

---

## Task Pane Debug Label

The task pane should show a compact development label when running from a local generated profile.

Example:

```text
Milton dev 8f3a21c4 | main | 1ea22c7 | :3184
```

Useful fields:

- short profile key
- Git branch
- short commit SHA
- dev server port
- build timestamp or Vite startup timestamp

This should be visually quiet, but always available in development. When two Excel task panes are open, the label should make it immediately clear which worktree each pane belongs to.

---

## Expected Developer Workflow

Create or open separate Codex threads, each with its own worktree.

Thread A:

```sh
pnpm dev
pnpm start
```

Thread B:

```sh
pnpm dev
pnpm start
```

Each thread gets:

- a different generated manifest
- a different manifest ID
- a different localhost port
- a different Excel ribbon entry/display name
- a different task pane debug label

Both add-ins can be loaded into Excel at the same time. Each task pane talks to the Vite server for its own worktree.

---

## Office Cache Behavior

Office can cache manifests, command surfaces, icons, and task pane web resources. The workflow should assume:

- JavaScript, CSS, and HTML changes usually only need a task pane reload.
- Manifest, ribbon, command, icon, or display-name changes may require stopping the sideloaded add-in and clearing Office cache.
- Generated manifests should change identity only when the worktree identity changes, not on every build.

Cache clearing should be treated as a debugging tool, not as the normal edit-refresh loop.

---

## Cleanup

There should be a cleanup command for generated local profile state.

Desired behavior:

```sh
pnpm dev:profiles
pnpm dev:clean-profile
pnpm dev:clean-stale-profiles
```

The cleanup command should:

- list known generated profiles
- show worktree path, port, and manifest path
- remove profiles whose worktree path no longer exists
- optionally stop/unload a selected generated manifest

The global registry should tolerate deleted worktrees. A stale entry should never prevent a new worktree from starting; it should only be reused if the recorded path still exists and matches the current worktree.

---

## Implementation Notes

The resolver should be a small Node/TypeScript utility used by all scripts. Avoid duplicating path hashing, port allocation, and manifest naming logic across package scripts.

Recommended module responsibilities:

```text
resolveDevProfile()
  -> worktreeRoot
  -> profileKey
  -> shortProfile
  -> displayName
  -> manifestId
  -> port
  -> manifestPath

generateDevManifest(profile)
  -> generated XML file
  -> generated metadata file

startDevServer(profile)
  -> Vite on profile.port

sideloadDevManifest(profile)
  -> office-addin-debugging start profile.manifestPath
```

The implementation should keep the release manifest path separate from generated development manifests. Production release behavior should not depend on local profile generation.
