/** Fresh-process route proof; no GC, memory instrumentation or module eviction. */
import { createRequire } from 'node:module';
import { IndexedWorkbook } from '../../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';
const require = createRequire(import.meta.url);
const providers = () =>
  Object.keys(require.cache).filter((path) => path.replaceAll('\\', '/').includes('/node_modules/@univerjs/'));
const before = providers();
let diagnostic;
const book = new IndexedWorkbook(
  native,
  () => {
    throw Error('Unexpected writer');
  },
  async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
);
try {
  await book.open(process.argv[2]);
  const cells = (await book.read('0', 'Cases', 'E1:H118')).cells;
  console.log(
    JSON.stringify({
      before,
      after: providers(),
      diagnostic,
      results: Object.fromEntries(
        Object.entries(cells)
          .filter(([, cell]) => cell.formula)
          .map(([ref, cell]) => [ref, cell.calculation])
      ),
    })
  );
} finally {
  await book.close();
  closeNativeWorkbookWorker();
}
