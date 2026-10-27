/** Build the component's native engine without the host application workspace. */
import { execFileSync } from 'node:child_process';
import { mkdirSync, copyFileSync, chmodSync, writeFileSync, readFileSync, renameSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import path from 'node:path';
import { createHash } from 'node:crypto';
const component = fileURLToPath(new URL('../', import.meta.url));
export function buildNativeWorkbook({ offline = process.env.CARGO_NET_OFFLINE === 'true' } = {}) {
  const manifest = path.join(component, 'native/Cargo.toml');
  execFileSync(
    'cargo',
    ['build', '--locked', ...(offline ? ['--offline'] : []), '--release', '--manifest-path', manifest],
    { stdio: 'inherit' }
  );
  const metadata = JSON.parse(
    execFileSync('cargo', ['metadata', '--no-deps', '--format-version', '1', '--manifest-path', manifest], {
      encoding: 'utf8',
    })
  );
  const executable = process.platform === 'win32' ? 'workbook-engine.exe' : 'workbook-engine';
  const source = path.join(metadata.target_directory, 'release', executable),
    directory = path.join(component, 'native/bin', `${process.platform}-${process.arch}`),
    binary = path.join(directory, executable);
  mkdirSync(directory, { recursive: true });
  // Install as a new file and swap it in. Overwriting the bytes of a program
  // that is running makes macOS refuse to start it again.
  const staged = binary + '.new-' + process.pid;
  copyFileSync(source, staged);
  renameSync(staged, binary);
  chmodSync(binary, 0o755);
  const identity = {
    schema: 'workbook.native.v1',
    platform: process.platform,
    arch: process.arch,
    sha256: createHash('sha256').update(readFileSync(binary)).digest('hex'),
    scope: 'Built for this host only; other platforms require separate qualification.',
  };
  writeFileSync(path.join(directory, 'identity.json'), JSON.stringify(identity, null, 2) + '\n');
  return { binary, directory, identity };
}
if (process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url) buildNativeWorkbook();
