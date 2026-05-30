# Milton
A Microsoft Excel Add-In for Bring-Your-Own-Model (BYOM) agentic spreadsheet interactions.

## Development

```sh
pnpm dev
pnpm start
```

`pnpm dev` starts the Office app's Vite HTTPS dev server on a worktree-derived local port. `pnpm start` generates and sideloads a worktree-specific Excel manifest so multiple Codex worktrees can be debugged side by side.

For the first agent-chat milestone, create a shared local dev env once:

```sh
pnpm dev:env-init
```

This creates `~/.config/milton/office-dev.env`. Values in that file are loaded by `pnpm dev`, `pnpm start`, and `pnpm validate` from every worktree.

The env file should contain:

```sh
DEBUG_OPENAI_API_KEY=sk-...
DEBUG_OPENAI_MODEL=gpt-5-mini
```

You can still create `apps/office/.env.local` inside one worktree for a worktree-specific override.

The task pane calls OpenAI directly from the browser runtime. This is a local development path only because `DEBUG_OPENAI_API_KEY` is exposed to the task pane.

Stop a sideloaded debug session with:

```sh
pnpm stop
```

## Project layout

```text
apps/
  office/
    commands/          Office command function page
    manifests/         Office manifests
    public/assets/     Static icon and logo assets served by Vite
    taskpanes/excel/   Excel task pane app
    vite.config.ts     Office app build config
docs/
packages/
  office-host/         Typed Office host integrations
  office-runtime/      OfficeJS code execution runtime and agent tool
  pi-agent-core/       Vendored Pi agent package, with a browser entrypoint for Milton
  pi-ai/               Vendored Pi model/provider abstraction package
  ui/                  Shared React UI
```

The source local Excel manifest is `apps/office/manifests/excel.local.xml`. Development sideloading uses generated manifests under `apps/office/.generated/`.

## Checks

```sh
pnpm build
pnpm validate
```

## Dependency notes

`pnpm-workspace.yaml` patches `office-addin-dev-settings` with `semver` because that package currently imports it at runtime without declaring it as a dependency.
