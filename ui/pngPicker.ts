import type { WorkbookPngAsset } from '../core/contracts';

/** Browser file capability yields bounded inline bytes; no path enters a command. */
export function pickWorkbookPng(): Promise<WorkbookPngAsset | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/png';
    input.style.display = 'none';
    const finish = (value: WorkbookPngAsset | null, error?: unknown): void => {
      input.remove();
      if (error) reject(error);
      else resolve(value);
    };
    input.oncancel = () => finish(null);
    input.onchange = () => {
      void (async () => {
        const file = input.files?.[0];
        if (!file) {
          finish(null);
          return;
        }
        if (file.size > 128 * 1024 || file.size < 8) throw Error('Choose a PNG image of at most 128 KiB.');
        const bytes = new Uint8Array(await file.arrayBuffer());
        if (![137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value))
          throw Error('Choose a PNG image.');
        const digest = await crypto.subtle.digest('SHA-256', bytes),
          sha256 = Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, '0')).join('');
        let binary = '';
        for (let offset = 0; offset < bytes.length; offset += 8192)
          binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
        finish({ mime: 'image/png', base64: btoa(binary), sha256, bytes: bytes.length });
      })().catch((error) => finish(null, error));
    };
    document.body.append(input);
    input.click();
  });
}
