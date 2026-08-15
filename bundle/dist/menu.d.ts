export interface MenuSuggestion {
    id: string;
    title: string;
    description: string;
    example: string;
}
export interface WorkbookMenu {
    summary: string;
    suggestions: MenuSuggestion[];
    note: string;
}
/**
 * Turn a workbook into a plain-language capability menu: a one-line digest of
 * what the file contains plus concrete next steps, each with an example prompt
 * the user can just confirm. This is the "let the agent speak first" entry
 * point for users who cannot describe what they want.
 */
export declare function buildWorkbookMenu(path: string, sheet?: string): Promise<WorkbookMenu>;
//# sourceMappingURL=menu.d.ts.map