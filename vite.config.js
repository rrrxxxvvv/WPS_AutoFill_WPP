import { defineConfig } from 'vite'
import { copyFile } from 'wpsjs/vite_plugins'
import { readFileSync } from 'node:fs'

export default defineConfig({
  base: './',
  plugins: [
    copyFile({ src: 'manifest.xml', dest: 'manifest.xml' }),
    copyFile({ src: 'js', dest: 'js' }),
    copyFile({ src: 'assets', dest: 'assets' }),
    copyFile({ src: 'LICENSE', dest: 'LICENSE' }),
    {
      name: 'third-party-license',
      generateBundle() {
        this.emitFile({ type: 'asset', fileName: 'LICENSE-THIRD-PARTY',
          source: readFileSync('THIRD_PARTY_NOTICES.md', 'utf8') })
      }
    },
    copyFile({ src: 'images', dest: 'images' }),
    copyFile({ src: 'ui', dest: 'ui' }),
    copyFile({ src: 'main.js', dest: 'main.js' }),
    copyFile({ src: 'ribbon.xml', dest: 'ribbon.xml' })
  ],
  server: { host: '0.0.0.0' }
})
