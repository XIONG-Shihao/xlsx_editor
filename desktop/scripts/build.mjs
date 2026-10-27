import {build} from 'esbuild'
import {build as buildFrontend} from 'vite'
import {mkdir,writeFile,realpath} from 'node:fs/promises'
import {resolve,join,relative} from 'node:path'
import {productRoot} from './tooling/stage.mjs'
const root=join(productRoot,'desktop'),editor=await realpath(productRoot),editorMain=join(editor,'standalone/main.mjs')
await mkdir(join(root,'.build'),{recursive:true});await mkdir(join(root,'out/main'),{recursive:true})
await writeFile(join(root,'out/main/package.json'),'{"type":"module"}\n')
await build({entryPoints:[join(root,'backend/index.ts')],bundle:true,platform:'node',format:'esm',target:'node24',outfile:join(root,'out/main/index.js'),packages:'external',plugins:[{name:'standalone-editor-files',setup(builder){
 // The editor's standalone main is bundled; every other editor file is loaded
 // from the editor files the package ships beside it.
 builder.onResolve({filter:/^\./},async args=>{
  const target=await realpath(resolve(args.resolveDir,args.path)).catch(()=>resolve(args.resolveDir,args.path))
  if(target.startsWith(editor+'/')&&target!==editorMain){
   const logical=join(productRoot,relative(editor,target))
   return {path:relative(join(root,'out/main'),logical).replaceAll('\\','/'),external:true}
  }
 })
}}],banner:{js:"import {createRequire as __nodeRequire} from 'node:module'; const require=__nodeRequire(import.meta.url); if(process.env.XLSX_EDITOR_RESOURCES)Object.defineProperty(process,'resourcesPath',{value:process.env.XLSX_EDITOR_RESOURCES});"}})
await buildFrontend({configFile:join(root,'vite.config.ts')})
