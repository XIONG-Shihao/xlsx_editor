export declare const REFUSAL_CODES: Readonly<Record<string, string>>;
/** An Error with a stable code from the registry. */
export declare function refusal(code: string, message: string): Error & { code: string };
/** True for one of our refusals (its code is in the registry), false for a system error such as ENOENT. */
export declare const isRefusal: (error: unknown) => boolean;
/** Re-word an error while keeping its code, for layers that add context. */
export declare function reworded(error: unknown, message: string, code?: string): Error & { code?: string };
