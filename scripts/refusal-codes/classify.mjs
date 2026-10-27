// Applies mapping.mjs to a site.
import { SIGNALS, EXCEPTIONS, KIND_RULES, FILE_DOMAINS, MESSAGE_DOMAINS, DEFAULT_DOMAIN } from './mapping.mjs';

/** Planning code prefixes its sentences (`validation: …`, `PRINT_X: …`); the rest of the sentence is what we classify. */
export const plain = (message) => message.replace(/^(validation|conflict|PRINT_[A-Z_]+): /, '');

export function kindOf(message) {
  message = plain(message);
  for (const [kind, pattern] of KIND_RULES) if (pattern.test(message)) return kind;
  return null;
}

export function domainOf(file, message) {
  message = plain(message);
  const name = file.replace(/^(native\/src|core|adapters)\//, '');
  for (const [pattern, domain] of FILE_DOMAINS) if (pattern.test(name)) return domain;
  for (const [pattern, domain] of MESSAGE_DOMAINS) if (pattern.test(message)) return domain;
  return DEFAULT_DOMAIN;
}

export const isSignal = (file, message) => SIGNALS.some((rule) => rule.file.test(file) && rule.message.test(message));

/** The code for a site, or null when no rule covers it. */
export function codeFor(file, message) {
  for (const rule of EXCEPTIONS)
    if ((!rule.file || rule.file.test(file)) && rule.message.test(message)) return rule.code;
  const kind = kindOf(message) ?? 'refused';
  if (kind.startsWith('internal-') || kind.startsWith('operation-') || kind.startsWith('session-')) return kind;
  return `${domainOf(file, message)}-${kind}`;
}
