import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// "base" muss exakt dem Repository-Namen entsprechen, da GitHub Pages die App
// unter https://<user>.github.io/<repo>/ ausliefert, nicht unter der Domain-Wurzel.
export default defineConfig({
  base: "/packliste/",
  plugins: [react()],
});
