import type { Writable } from 'node:stream'
import { frame } from './framing'

/** Limits include the packet currently blocked on the pipe, not only waiters.
 * Saturation fails the connection: dropping a document/agent event would leave
 * a plausibly healthy UI with incomplete state. */
export class FrameWriter {
  private tail = Promise.resolve()
  private queuedBytes = 0
  private queuedFrames = 0
  private failure: Error | undefined
  private rejectWriting: ((error: Error) => void) | undefined
  constructor(private readonly stream: Writable, private readonly onFailure: (error: Error) => void,
    private readonly maxBytes = 64 * 1024 * 1024, private readonly maxFrames = 512) {
    stream.on('error', error => this.fail(error))
    stream.on('close', () => this.fail(Error('unavailable: Desktop pipe closed')))
  }
  private fail(error: Error): Error {
    if (!this.failure) {
      this.failure = error
      this.rejectWriting?.(error); this.rejectWriting = undefined
      this.onFailure(error)
    }
    return this.failure
  }
  send(value: unknown): Promise<void> {
    if (this.failure) return Promise.reject(this.failure)
    let packet: Buffer
    try { packet = frame(value) } catch (error) { return Promise.reject(error) }
    if (this.queuedBytes + packet.length > this.maxBytes || this.queuedFrames >= this.maxFrames) {
      return Promise.reject(this.fail(Error('unavailable: Desktop output budget exceeded; connection must restart')))
    }
    this.queuedBytes += packet.length; this.queuedFrames++
    const pending = this.tail.then(() => {
      if (this.failure) throw this.failure
      return new Promise<void>((resolve, reject) => {
        this.rejectWriting = reject
        // The callback completes only after the writable consumes this packet.
        this.stream.write(packet, error => {
          this.rejectWriting = undefined
          error ? reject(this.fail(error)) : resolve()
        })
      })
    }).finally(() => { this.queuedBytes -= packet.length; this.queuedFrames-- })
    this.tail = pending.catch(() => {})
    return pending
  }
}
