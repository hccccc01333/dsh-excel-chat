import type { Context } from '@deepseek-ai/cordis';
/** One registration that threw, kept so `apply` can summarise them at the end. */
export interface RegistrationFailure {
    label: string;
    detail: string;
}
export interface RegistrationReport {
    registered: string[];
    failures: RegistrationFailure[];
}
type LogLevel = 'info' | 'warn' | 'error';
/** Flatten anything thrown into one line that survives being logged. */
export declare function describeError(error: unknown): string;
/**
 * Write one line to every channel the host might actually persist.
 *
 * The host logger is the only sink that reaches its log file, but it is not
 * always mounted — and reading it can itself throw — so the console is kept
 * alongside it rather than replaced by it.
 */
export declare function announce(host: unknown, level: LogLevel, line: string): void;
/** One line describing what actually came up, so a partial load is never silent. */
export declare function registrationSummary(report: RegistrationReport): string;
/**
 * Wrap a host context so that one failing registration cannot take the rest of
 * the plugin down with it.
 *
 * Registration runs inside the host's own startup path, and cordis does not
 * isolate a throw there: `ctx.effect` rethrows whatever its callback threw, so
 * the throw escapes `apply`, and a failed `apply` disposes *every* effect the
 * plugin had registered. One malformed tool definition therefore removes all the
 * others too, and the plugin looks like it was never loaded at all — no tools, no
 * error, nothing in the log.
 *
 * Each effect is therefore run through a try/catch that reports the failure with
 * the label of the thing that failed, and returns normally so the remaining
 * registrations still happen.
 */
export declare function guardedContext(host: Context): {
    ctx: Context;
    report: RegistrationReport;
};
export {};
//# sourceMappingURL=registration.d.ts.map