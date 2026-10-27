import {readFile,writeFile,mkdir,cp} from 'node:fs/promises'
import {join} from 'node:path'
import {constants} from 'node:fs'
import {sourceBun} from './tooling/bun.mjs'
import {productRoot,run} from './tooling/stage.mjs'

export async function stageDependencies(app){
 const component=JSON.parse(await readFile(join(productRoot,'package.json'),'utf8'))
 const manifest={...component,private:true,scripts:undefined,files:undefined,exports:undefined}
 await mkdir(app,{recursive:true})
 await writeFile(join(app,'package.json'),JSON.stringify(manifest,null,2)+'\n')
 await cp(join(productRoot,'bun.lock'),join(app,'bun.lock'),{mode:constants.COPYFILE_FICLONE})
 await run(sourceBun,['install','--production','--frozen-lockfile','--ignore-scripts','--linker','hoisted'],{cwd:app})
}
