export interface StandalonePorts {
  invoke: (channel: string, ...args: unknown[]) => Promise<unknown>;
  subscribe: (channel: string, callback: (value: unknown) => void) => () => void;
}
export function createStandaloneApi(ports: StandalonePorts): Record<string, (...args: any[]) => unknown>;
