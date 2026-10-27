// OPC part names are package identifiers, not operating-system paths. These
// utilities deliberately expose neither filesystem access nor a Node path API.
import { refusal } from './refusal-codes.mjs';

function ownerSegments(owner) {
  if (
    typeof owner !== 'string' ||
    !owner ||
    owner.startsWith('/') ||
    owner.includes('\\') ||
    owner.split('/').some((part) => !part || part === '.' || part === '..')
  )
    throw refusal('internal-invalid-field', 'Invalid package owner');
  return owner.split('/');
}
export function relationshipPart(owner) {
  const parts = ownerSegments(owner),
    name = parts.pop();
  return [...parts, '_rels', name + '.rels'].join('/');
}
export function resolvePackageTarget(owner, target) {
  const parts = ownerSegments(owner);
  parts.pop();
  if (typeof target !== 'string' || !target || target.includes('\\') || /[?#]/.test(target))
    throw refusal('package-malformed', 'Invalid package relationship target');
  if (target.startsWith('/')) parts.length = 0;
  for (const part of target.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      if (!parts.length) throw refusal('package-unsupported', 'Package relationship escapes its root');
      parts.pop();
    } else parts.push(part);
  }
  if (!parts.length) throw refusal('internal-missing-field', 'Missing package target part');
  return parts.join('/');
}
