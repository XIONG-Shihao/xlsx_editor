import {defineConfig} from 'vite'
import react from '@vitejs/plugin-react'
import {readdirSync} from 'node:fs'
import {resolve} from 'node:path'
export default defineConfig({
  base:'./', plugins:[react()], worker:{format:'es'},
  build:{outDir:'dist/ui',emptyOutDir:true,lib:{entry:Object.fromEntries(readdirSync('ui').filter(n=>/\.(ts|tsx)$/.test(n)&&!n.endsWith('.worker.ts')).map(n=>[n.replace(/\.(ts|tsx)$/,''),resolve('ui',n)])),formats:['es'],fileName:(_format,name)=>name+'.js'},rollupOptions:{external:id=>/^react(?:-dom)?(?:\/|$)/.test(id),output:{chunkFileNames:'chunks/[name]-[hash].js'}}}
})
