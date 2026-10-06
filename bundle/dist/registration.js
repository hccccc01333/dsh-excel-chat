import { writeFileSync } from 'node:fs';
/** Flatten anything thrown into one line that survives being logged. */
export function describeError(error) {
    if (error instanceof Error)
        return error.stack ?? `${error.name}: ${error.message}`;
    return String(error);
}
/**
 * Write one line to every channel the host might actually persist.
 *
 * The host logger is the only sink that reaches its log file, but it is not
 * always mounted — and reading it can itself throw — so the console is kept
 * alongside it rather than replaced by it.
 */
export function announce(host, level, line) {
    const sinks = [];
    try {
        const logger = host?.logger;
        const method = logger?.[level];
        if (typeof method === 'function') {
            sinks.push((text) => method.call(logger, text));
        }
    }
    catch {
        // A logger getter that throws must not swallow the message we are writing.
    }
    sinks.push(level === 'info' ? console.log : level === 'warn' ? console.warn : console.error);
    for (const sink of sinks) {
        try {
            sink(line);
        }
        catch {
            // Reporting must never be the thing that throws.
        }
    }
}
/** One line describing what actually came up, so a partial load is never silent. */
export function registrationSummary(report) {
    if (report.failures.length === 0) {
        return `[dsh-excel-chat] ready: ${report.registered.length} tools registered`;
    }
    const failed = report.failures.map((failure) => failure.label).join(', ');
    return `[dsh-excel-chat] PARTIAL: ${report.registered.length} tools registered, ${report.failures.length} failed -> ${failed}`;
}
/**
 * Write the registration outcome to `DSH_EXCEL_CHAT_STATUS`, when that is set.
 *
 * The host owns `apply()`, so when a host never brings the plugin up there is no
 * channel left to ask — no tools, no error, nothing in the log. A status file
 * settles the one question that separates the two failure modes: a file means
 * `apply` ran, and names every tool it managed to register; no file means the
 * row never mounted at all. Off unless the variable is set, so a normal load
 * touches no disk.
 */
export function writeStatusReport(report, target = process.env.DSH_EXCEL_CHAT_STATUS) {
    if (target === undefined || target === '')
        return;
    try {
        writeFileSync(target, JSON.stringify({
            applied: true,
            tools: report.registered,
            failures: report.failures.map((failure) => ({
                label: failure.label,
                detail: failure.detail.split('\n')[0],
            })),
            at: new Date().toISOString(),
        }, undefined, 2) + '\n');
    }
    catch {
        // Diagnostics must never be the reason a plugin fails to load.
    }
}
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
export function guardedContext(host) {
    const report = { registered: [], failures: [] };
    const fail = (label, error) => {
        const detail = describeError(error);
        report.failures.push({ label, detail });
        announce(host, 'error', `[dsh-excel-chat] ${label} failed to register: ${detail}`);
    };
    const ctx = {
        get: (key) => host.get(key),
        effect: (callback, label) => host.effect(() => {
            try {
                // The disposer a registration returns is forwarded untouched; cordis
                // validates the shape itself, so it stays opaque here.
                return callback();
            }
            catch (error) {
                fail(label ?? 'an unnamed effect', error);
                // Returning nothing (rather than rethrowing) is the whole point: the
                // remaining registrations still run.
                return undefined;
            }
        }, label),
        tools: {
            register: (definition) => {
                const name = definition?.name;
                const disposer = host.tools.register(definition);
                // Counted only after the host accepted it, so the summary never claims a
                // tool that failed.
                report.registered.push(typeof name === 'string' ? name : '<unnamed tool>');
                return disposer;
            },
        },
        systemPrompt: {
            section: (section) => host.systemPrompt.section(section),
        },
    };
    return { ctx: ctx, report };
}
