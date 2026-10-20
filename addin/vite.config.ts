import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

// Office add-ins must be served over HTTPS. `npm run dev` uses the trusted
// localhost certificate from office-addin-dev-certs when it is installed.
async function httpsOptions() {
  if (process.env.VITEST) return undefined;
  try {
    const certs = await import("office-addin-dev-certs");
    return await certs.getHttpsServerOptions();
  } catch {
    return undefined;
  }
}

export default defineConfig(async () => ({
  // GitHub Pages serves the add-in from /auditaitool/; local dev serves from /.
  base: process.env.BASE_PATH ?? "/",
  plugins: [react()],
  server: { port: 3000, https: await httpsOptions() },
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html"),
        viewer: resolve(__dirname, "viewer.html"),
      },
    },
  },
  test: { environment: "node" },
}));
