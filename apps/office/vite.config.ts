import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react-swc";
import { defineConfig } from "vite";
import * as devCerts from "office-addin-dev-certs";

const rootDir = dirname(fileURLToPath(import.meta.url));

export default defineConfig(async ({ command }) => {
  const server =
    command === "serve"
      ? {
          host: "localhost",
          port: 3000,
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
