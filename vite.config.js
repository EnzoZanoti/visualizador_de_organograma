import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // Importante para o GitHub Pages achar os arquivos corretos:
  base: '/visualizador_de_organograma/', 
})