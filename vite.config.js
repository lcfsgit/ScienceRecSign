import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  plugins: [vue()],
  base: process.env.VITE_BASE || '/',
  server: {
    host: '127.0.0.1',
    port: 5173
  },
  optimizeDeps: {
    include: ['xlsx', 'jszip']
  }
})
