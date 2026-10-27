import ts from 'typescript';
import { readFileSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { mkdirSync, copyFileSync, readdirSync } from 'node:fs';
execFileSync('npx', ['--no-install', 'tsc', '-p', 'tsconfig.build.json'], { stdio: 'inherit' });
mkdirSync('dist/types/core', { recursive: true });
for (const name of readdirSync('core'))
  if (name.endsWith('.d.mts')) copyFileSync('core/' + name, 'dist/types/core/' + name);
mkdirSync('dist/types/adapters', { recursive: true });
for (const name of readdirSync('adapters'))
  if (name.endsWith('.d.mts')) copyFileSync('adapters/' + name, 'dist/types/adapters/' + name);

mkdirSync('dist/core', { recursive: true });
for (const name of readdirSync('core'))
  if (name.endsWith('.ts') && !name.endsWith('.d.ts')) {
    const result = ts.transpileModule(readFileSync('core/' + name, 'utf8'), {
      compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ES2022 },
    });
    const output = result.outputText;
    writeFileSync('dist/core/' + name.replace(/\.ts$/, '.js'), output);
    // A plain module one of these files imports (the refusal-code list) must sit beside it.
    for (const [, module] of output.matchAll(/from '\.\/([\w-]+\.mjs)'/g))
      copyFileSync('core/' + module, 'dist/core/' + module);
  }
