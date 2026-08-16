export interface DoctorCheck {
    name: string;
    ok: boolean;
    detail: string;
}
/** Host packages that must stay peer/dev-only; a profile installing them as
 * dependencies shadows the harness host and breaks tool dispatch (Issue #1). */
export declare const HOST_PACKAGES: string[];
/**
 * Install self-check (doctor): environment, host-package isolation, and a
 * real engine smoke test that creates a temp workbook and reads/validates it.
 * The smoke test intentionally bypasses the harness scheduler, so it works
 * even when host Symbol identity is broken.
 */
export declare function runDoctorChecks(options?: {
    profileDirs?: string[];
}): Promise<DoctorCheck[]>;
//# sourceMappingURL=doctor.d.ts.map