# Milton
A Microsoft Excel Add-In for Bring-Your-Own-Model (BYOM) agentic spreadsheet interactions.

## Development

```sh
pnpm dev
pnpm start
```

`pnpm dev` starts the Office app's Vite HTTPS dev server on port 3000. `pnpm start` sideloads `apps/office/manifests/excel.local.xml` into Excel.

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
  ui/                  Shared React UI
```

The active Excel manifest is `apps/office/manifests/excel.local.xml`.

## Checks

```sh
pnpm build
pnpm validate
```

## Dependency notes

`pnpm-workspace.yaml` patches `office-addin-dev-settings` with `semver` because that package currently imports it at runtime without declaring it as a dependency.
