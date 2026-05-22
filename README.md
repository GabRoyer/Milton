# Milton
A Microsoft Excel Add-In for Bring-Your-Own-Model (BYOM) agentic spreadsheet interactions.

## Development

```sh
pnpm dev
pnpm start
```

`pnpm dev` starts the Vite HTTPS dev server on port 3000. `pnpm start` sideloads `office/manifest.xml` into Excel.

Stop a sideloaded debug session with:

```sh
pnpm stop
```

## Project layout

```text
app/
  commands/      Office command function page
  taskpane/      React task pane app
office/          Office manifest
public/assets/   Static icon and logo assets served by Vite
```

The active Excel manifest is `office/manifest.xml`.

## Checks

```sh
pnpm build
pnpm validate
```

## Dependency notes

`semver` is listed explicitly because `office-addin-dev-settings` currently imports it at runtime without declaring it as a dependency.
