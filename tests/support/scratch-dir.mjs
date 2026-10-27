// Creates .build/tests in the product folder, the ignored scratch folder tests write into.
// Import it before anything that makes a folder under it.
import { mkdirSync } from 'node:fs';

mkdirSync(new URL('../../.build/tests/', import.meta.url), { recursive: true });
