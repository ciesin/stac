import { defineConfig } from "vite";

export default defineConfig({
  resolve: {
    // app.js mounts an in-DOM template (index.html), which needs Vue's
    // runtime+compiler build — the default "vue" alias is runtime-only.
    alias: {
      vue: "vue/dist/vue.esm-bundler.js",
    },
  },
});
