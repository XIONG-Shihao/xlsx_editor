// The XLSX editor, installed from the product's own lockfile and built.
import {sourceBun,prepareBun} from './tooling/bun.mjs'
import {run} from './tooling/stage.mjs'
await prepareBun()
await run(sourceBun,['install','--frozen-lockfile','--ignore-scripts','--linker','hoisted'])
await run('npm',['run','build'])
