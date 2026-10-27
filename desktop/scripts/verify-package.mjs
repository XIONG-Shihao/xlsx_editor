import {verifyBunDependencies} from './tooling/bun-lock.mjs'
import {bunVersion,verifyBun} from './tooling/bun.mjs'
import {files,hash,natives,signNativeResources} from './tooling/native-resources.mjs'
import {productRoot} from './tooling/stage.mjs'
export {signNativeResources}
import {realpath,readFile,stat,access} from 'node:fs/promises'
import {join,relative,resolve} from 'node:path'
import {execFileSync} from 'node:child_process'
import {fileURLToPath} from 'node:url'
export async function verifyPackage(bundle){
 bundle=await realpath(resolve(bundle));const root=join(bundle,'Contents/Resources'),paths=await files(bundle,bundle)
 for(const path of paths){const name=relative(bundle,path);if(/(^|\/)(electron|codex|uv|python|superdoc|officemcp)(\/|$)/i.test(name)||/Electron Framework|\.app\//.test(name))throw Error('Unexpected integrated runtime: '+name)}
 for(const required of ['app/desktop/out/main/index.js','app/adapters/owner-worker.mjs','app/standalone/recent-files.mjs','backend/bun','native/workbook-engine','native/identity.json'])if(!(await stat(join(root,required))).isFile())throw Error('Missing standalone resource: '+required)
 const manifest=JSON.parse(await readFile(join(root,'standalone-package.json'),'utf8'));if(manifest.product!=='standalone-xlsx-tauri'||manifest.channel!=='internal')throw Error('Wrong standalone profile')
 const info=JSON.parse(execFileSync('plutil',['-convert','json','-o','-',join(bundle,'Contents/Info.plist')],{encoding:'utf8'}))
 const association=info.CFBundleDocumentTypes?.find(item=>item.CFBundleTypeName==='Spreadsheet')
 if(association?.CFBundleTypeRole!=='Editor'||association?.LSHandlerRank!=='Alternate'||JSON.stringify(association?.CFBundleTypeExtensions)!==JSON.stringify(['xlsx','xlsm','xls','xlsb','csv','tsv']))throw Error('Spreadsheet associations are missing or change the default handler rank')
 const pathsNative=await natives(bundle);if(pathsNative.length!==manifest.native.length)throw Error('Native inventory differs')
 for(const entry of manifest.native){const path=join(root,entry.path);if(relative(root,resolve(path)).startsWith('..')||!pathsNative.includes(path)||await hash(path)!==entry.sha256)throw Error('Invalid native resource inventory');execFileSync('codesign',['--verify','--strict',path],{stdio:'pipe'});const linked=execFileSync('otool',['-L',path],{encoding:'utf8'});if(/\/opt\/homebrew|\/usr\/local|Electron Framework/.test(linked))throw Error('Nonportable linkage: '+entry.path)}
 const coordinator=await readFile(join(root,'app/desktop/out/main/index.js'),'utf8');if(/(?:from\s*|require\s*\(|import\s*\()\s*['"]electron['"]/.test(coordinator))throw Error('Electron import survived in coordinator')
 if(manifest.node!==null || manifest.javascriptRuntime?.kind!=='bun' || manifest.javascriptRuntime.version!==bunVersion || manifest.javascriptRuntime.executable!=='backend/bun')throw Error('Wrong standalone runtime manifest')
 for(const file of ['backend/node','backend/node-provenance.json']){try{await access(join(root,file));throw Error('Unexpected Node payload: '+file)}catch(error){if(error.code!=='ENOENT')throw error}}
 await verifyBun(join(root,'backend/bun'),join(root,'backend/bun-provenance.json'))
 verifyBunDependencies(join(root,'app'),join(root,'backend/bun'))
 const bun=join(root,'backend/bun');if(execFileSync(bun,['--version'],{encoding:'utf8',env:{PATH:'/usr/bin:/bin'}}).trim()!==bunVersion)throw Error('Wrong portable Bun version')
 if(execFileSync(join(bundle,'Contents/MacOS/xlsx-desktop'),['--print-host-runtime'],{encoding:'utf8',env:{PATH:'/usr/bin:/bin'}}).trim()!==JSON.stringify({kind:'bun'}))throw Error('Native standalone launcher does not use Bun')
 execFileSync('codesign',['--verify','--deep','--strict',bundle],{stdio:'pipe'})
 console.log(JSON.stringify({verified:true,product:manifest.product,files:paths.length,native:pathsNative.length}))
}
if(process.argv[1]&&resolve(process.argv[1])===fileURLToPath(import.meta.url))await verifyPackage(process.argv[2]??join(productRoot,'desktop/release/mac-arm64/internal_release/XLSX Editor.app'))
