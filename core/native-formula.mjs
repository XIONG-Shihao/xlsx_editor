/** The native subset is an optimization, never a new workbook acceptance limit.
 * A miss leaves the unpublished projection unchanged for the compatibility engine. */

export async function nativeFormula(native, root, method, payload) {
  const request = { action: 'workbook', operation: 'formula', root, method, ...payload };
  let nonfinite = false;
  const wire = JSON.stringify(request, (_key, value) => {
    if (typeof value === 'number' && (!Number.isFinite(value) || Object.is(value, -0))) nonfinite = true;
    return value;
  });
  if (nonfinite) return { supported: false, reason: 'Nonfinite native input' };
  if (Buffer.byteLength(wire) > 2 * 1024 * 1024) return { supported: false, reason: 'Native projection byte bound' };
  return native(request);
}
