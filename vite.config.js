import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  base: "/sanjog/",   // ← must match your GitHub repo name exactly
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["icon.png"],
      manifest: {
        name: "Sanjog",
        short_name: "Sanjog",
        description: "Your personal recruitment platform advisor",
        theme_color: "#E8930A",
        background_color: "#0F0E0C",
        display: "standalone",
        orientation: "portrait",
        start_url: "/sanjog/",
        scope: "/sanjog/",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png", purpose: "any maskable" }
        ]
      },
      workbox: {
        globPatterns: ["**/*.{js,css,html,ico,png,svg,woff2}"],
        runtimeCaching: [{
          urlPattern: /^https:\/\/api\.anthropic\.com\/.*/i,
          handler: "NetworkOnly"
        }]
      }
    })
  ]
});
