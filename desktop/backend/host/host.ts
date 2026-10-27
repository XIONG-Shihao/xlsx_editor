import { Frames } from './framing'
import { FrameWriter } from './frameWriter'

/** The standalone shell's connection to its Rust host: framed stdin/stdout.
 * Requests come in as `invoke`, native actions go out as `native` and come
 * back as `native-result`, and events for the window go out as `event`. */
let failed: Error | undefined
let onFailure: (error: Error) => void = console.error
function failTransport(error: Error): void {
  if (failed) return
  failed = error
  for (const pending of actions.values()) { clearTimeout(pending.timer); pending.reject(error) }
  actions.clear()
  onFailure(error)
}
const output = new FrameWriter(process.stdout, failTransport)
export const send = (value: unknown): Promise<void> => output.send(value)
const actions = new Map<string, { resolve: (value: unknown) => void; reject: (error: Error) => void; timer: ReturnType<typeof setTimeout> }>()
let sequence = 0

/** A window, dialog or system action performed by the Rust host. */
export function hostAction(action: string, options?: unknown): Promise<unknown> {
  return new Promise((resolve, reject) => {
    if (failed) { reject(failed); return }
    if (actions.size >= 32) { reject(Error('busy: Too many native operations')); return }
    const id = `native-${++sequence}`
    const timer = setTimeout(() => { actions.delete(id); reject(Error('timeout: Native operation timed out')) }, 120000)
    actions.set(id, { resolve, reject, timer })
    void send({ kind: 'native', id, action, options }).catch(error => { clearTimeout(timer); actions.delete(id); reject(error) })
  })
}
const last = (values: unknown[]): unknown => values.at(-1)
export const dialog = {
  showOpenDialog: (...values: unknown[]) => hostAction('openDialog', last(values)) as Promise<{ canceled: boolean; filePaths: string[] }>,
  showSaveDialog: (...values: unknown[]) => hostAction('saveDialog', last(values)) as Promise<{ canceled: boolean; filePath?: string }>,
  showMessageBox: (...values: unknown[]) => hostAction('messageDialog', last(values)) as Promise<{ response: number }>
}
export const clipboard = {
  writeText: (text: string) => hostAction('clipboardWrite', { text }),
  readText: () => hostAction('clipboardRead') as Promise<string>
}

export function setupHost(handle: (channel: string, args: unknown[]) => Promise<unknown>, failure: (error: Error) => void): void {
  onFailure = failure
  // stdout is framed protocol only. Never let dependency logs corrupt it.
  console.log = console.info = console.debug = (...args) => console.error(...args)
  const input = process.stdin.pipe(new Frames())
  let active = 0
  input.on('data', message => {
    if (!message || typeof message !== 'object') { failTransport(Error('invalid_request: Invalid desktop envelope')); return }
    if (message.kind === 'transport-stop') { failTransport(Error(String(message.error ?? 'unavailable: Desktop transport stopped'))); return }
    if (message.kind === 'native-result') {
      const pending = actions.get(message.id)
      if (pending) { clearTimeout(pending.timer); actions.delete(message.id); message.error ? pending.reject(Error(message.error)) : pending.resolve(message.result) }
      return
    }
    if (message.kind !== 'invoke' || typeof message.id !== 'string' || typeof message.channel !== 'string' || !Array.isArray(message.args)) {
      void send({ id: message.id, error: 'invalid_request: Invalid desktop request' }).catch(console.error); return
    }
    if (active >= 512) { void send({ id: message.id, error: 'busy: Too many desktop operations' }).catch(console.error); return }
    active++
    void handle(message.channel, message.args).then(result => send({ id: message.id, result }).catch(error => {
      if (failed) throw error
      // A response exceeding one frame is an explicit operation failure, not a silent timeout.
      return send({ id: message.id, error: String(error?.message ?? error) })
    }), error => send({ id: message.id, error: String(error?.message ?? error) }))
      .catch(console.error).finally(() => { active-- })
  })
  input.on('error', failTransport)
}
