import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'

const base = '/HSG-Game-Manager/'

export default defineConfig({
  base,

  plugins: [
    react(),

    VitePWA({
      registerType: 'autoUpdate',

      includeAssets: [
        'favicon.svg',
      ],

      manifest: {
        name: 'HSG Game Manager',
        short_name: 'HSG Manager',

        description:
          'Hunting Serious Game management and wildlife monitoring',

        start_url: base,
        scope: base,

        display: 'standalone',

        background_color: '#ffffff',
        theme_color: '#176447',

        icons: [
          {
            src: 'icon-192.png',
            sizes: '192x192',
            type: 'image/png',
          },
          {
            src: 'icon-512.png',
            sizes: '512x512',
            type: 'image/png',
          },
        ],
      },

      workbox: {
        globPatterns: [
          '**/*.{js,css,html,svg,png,ico,json}',
        ],

        maximumFileSizeToCacheInBytes:
          10 * 1024 * 1024,
      },
    }),
  ],
})