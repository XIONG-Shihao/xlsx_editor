import {defineConfig} from 'vite'
import react from '@vitejs/plugin-react'
import {dirname} from 'node:path'
import {fileURLToPath} from 'node:url'
import {resolve} from 'node:path'
const root=dirname(fileURLToPath(import.meta.url))
export default defineConfig({root:resolve(root,'frontend'),base:'./',plugins:[react()],worker:{format:'es'},resolve:{dedupe:['react','react-dom']},build:{outDir:resolve(root,'out/renderer'),emptyOutDir:true,target:'safari17'}})
