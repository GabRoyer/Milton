import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react-swc";
import { defineConfig } from "vite";
import * as devCerts from "office-addin-dev-certs";

const rootDir = dirname(fileURLToPath(import.meta.url));
const devServerPort = Number(process.env.MILTON_OFFICE_PORT ?? 3000);

export default defineConfig(async ({ command }) => {
  const server =
    command === "serve"
      ? {
          host: "localhost",
          port: devServerPort,
          strictPort: true,
          https: await devCerts.getHttpsServerOptions(),
          headers: {
            "Access-Control-Allow-Origin": "*",
          },
        }
      : undefined;

  return {
    plugins: [react()],
    publicDir: "public",
    server,
    envPrefix: ["DEBUG_", "MILTON_PUBLIC_"],
    build: {
      outDir: "dist",
      rollupOptions: {
        input: {
          taskpane: resolve(rootDir, "taskpanes/excel/taskpane.html"),
          commands: resolve(rootDir, "commands/commands.html"),
        },
      },
    },
  };
});
