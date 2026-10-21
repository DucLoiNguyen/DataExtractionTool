import { fileURLToPath, URL } from 'node:url'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { defineConfig } from 'vitest/config'

// Khi chay dev (npm run dev) trinh duyet goi /api -> chuyen sang backend. Khi trien khai that
// thi lop proxy (Nginx...) o phia truoc dam nhiem viec nay, khong dung toi cau hinh nay.
const apiTarget = process.env.VITE_API_PROXY ?? 'http://127.0.0.1:8000'

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: { alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) } },
  server: { port: 5173, proxy: { '/api': apiTarget } },
  test: { environment: 'node', include: ['src/**/*.test.ts'] },
})
