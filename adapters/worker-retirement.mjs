/** Register immediately after construction, before a worker can exit.
 * Exit is authoritative. Some hosts never resolve terminate() called again on
 * an exited worker; do not let that strand document cleanup or recovery leases.
 */
export function workerRetirement(worker) {
  let exited = false,
    retirement;
  const exit = new Promise((resolve) =>
    worker.once('exit', () => {
      exited = true;
      resolve();
    })
  );
  return () => (retirement ??= exited ? exit : Promise.race([worker.terminate(), exit]).then(() => exit));
}
