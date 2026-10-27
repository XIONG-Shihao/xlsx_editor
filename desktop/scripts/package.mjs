import {stageDependencies} from './stage-dependencies.mjs'
import {bunManifest,prepareBun,sourceBun} from './tooling/bun.mjs'
// Independent internal product. Overwrite after verification; never retain
// history/ versions. This command does not build a public product or full app.
import {cp,writeFile,mkdir,rm} from 'node:fs/promises'
import {constants} from 'node:fs'
import {resolve,join} from 'node:path'
import {execFileSync} from 'node:child_process'
import {productRoot,run,publishBundle} from './tooling/stage.mjs'
import {verifyPackage,signNativeResources} from './verify-package.mjs'
if(process.platform!=='darwin'||process.arch!=='arm64')throw Error('Standalone packaging targets macOS arm64')
if(process.argv.length>2)throw Error('Only the regular internal standalone product is implemented')
const root=join(productRoot,'desktop'),stage=join(root,'.build/stage'),app=join(stage,'app'),resources=join(stage,'backend')
await prepareBun()
await run(process.execPath,['desktop/scripts/build-editor.mjs'])
await run(process.execPath,['desktop/scripts/build.mjs'])
await rm(stage,{recursive:true,force:true});await mkdir(app,{recursive:true});await mkdir(resources,{recursive:true})
const copy=(source,destination)=>cp(source,destination,{recursive:true,mode:constants.COPYFILE_FICLONE})
await stageDependencies(app)
// The app folder mirrors the product folder, so the coordinator's relative paths hold.
for(const entry of ['core','adapters','dist/core','standalone/recent-files.mjs','standalone/platform.mjs'])await copy(join(productRoot,entry),join(app,entry))
await copy(join(root,'out/main'),join(app,'desktop/out/main'))
await copy(sourceBun,join(resources,'bun'))
await copy(join(sourceBun,'../../manifest.json'),join(resources,'bun-provenance.json'))
await copy(join(productRoot,'native/bin/darwin-arm64'),join(stage,'native'))
const config={build:{beforeBuildCommand:''},bundle:{active:true,targets:['app'],resources:{'../.build/stage/app/':'app/','../.build/stage/backend/':'backend/','../.build/stage/native/':'native/'},macOS:{signingIdentity:'-'}}}
const configFile=join(root,'.build/package-config.json');await writeFile(configFile,JSON.stringify(config,null,2))
await run(join(productRoot,'node_modules/.bin/tauri'),['build','--bundles','app','--config',configFile],{cwd:root})
const bundle=join(root,'src-tauri/target/release/bundle/macos/XLSX Editor.app'),bundleResources=join(bundle,'Contents/Resources')
const native=await signNativeResources(bundle)
await writeFile(join(bundleResources,'standalone-package.json'),JSON.stringify({product:'standalone-xlsx-tauri',status:'migration-candidate',channel:'internal',node:null,javascriptRuntime:bunManifest('backend/bun'),native,components:['xlsx'],excluded:['full-app','codex','python','superdoc','officemcp','electron']},null,2)+'\n')
execFileSync('codesign',['--force','--sign','-',bundle],{stdio:'pipe'})
await verifyPackage(bundle)
const output=join(root,'release/mac-arm64/internal_release'),destination=join(output,'XLSX Editor.app')
const checkIdle=()=>{const processes=execFileSync('ps',['-axo','command'],{encoding:'utf8'});if(processes.split('\n').some(line=>line.startsWith(join(destination,'Contents/MacOS')+'/')))throw Error('Quit the destination standalone editor before replacement')}
await mkdir(output,{recursive:true});await publishBundle(bundle,destination,verifyPackage,checkIdle)
await rm(stage,{recursive:true,force:true});await rm(bundle,{recursive:true,force:true})
console.log('Standalone internal migration candidate: '+destination)
