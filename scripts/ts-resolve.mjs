/**
 * Resolver hook so Node's built-in type stripping can load the editor's `.ts` sources
 * with the project's extensionless relative imports (Vite resolves those at
 * build time; plain Node does not).
 */
export function resolve(specifier, context, next) {
  if (specifier.startsWith('.') && !/\.(ts|mjs|cjs|js|json|node)$/.test(specifier)) {
    try {
      return next(`${specifier}.ts`, context);
    } catch {
      /* fall through to the default resolution */
    }
  }
  return next(specifier, context);
}
