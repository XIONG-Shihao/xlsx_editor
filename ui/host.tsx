import type { ComponentType } from 'react';

/** Services supplied by an embedding application; no account or DOCX dependency. */
export interface XlsxHost {
  api: { copyText(text: string): Promise<void>; readClipboardText(): Promise<string> };
  Icon: ComponentType<{ name: string; size?: number; className?: string }>;
  useI18n(): { locale: 'en' | 'zh' };
}
let configured: XlsxHost | undefined;
export function configureXlsxHost(host: XlsxHost): void {
  configured = host;
}
export function editorHost(): XlsxHost {
  if (!configured) throw Error('Configure the XLSX host before rendering the editor');
  return configured;
}
export function Icon(props: { name: string; size?: number; className?: string }): JSX.Element {
  const Component = editorHost().Icon;
  return <Component {...props} />;
}
export function useI18n(): ReturnType<XlsxHost['useI18n']> {
  return editorHost().useI18n();
}
