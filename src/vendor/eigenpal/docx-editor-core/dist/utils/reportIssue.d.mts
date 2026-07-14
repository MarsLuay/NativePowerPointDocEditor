/**
 * Framework-agnostic helper to open a pre-filled GitHub "new issue"
 * URL for the Help > Report issue action. GitHub doesn't allow file
 * attachments via URL params, so the body asks the user to drag-drop
 * their DOCX onto the issue after it opens.
 *
 * Lifted from packages/react/src/components/reportIssue.ts so the
 * Vue adapter can re-use it without re-implementing.
 * @packageDocumentation
 * @public
 */
interface ReportIssueEnv {
    userAgent?: string;
    viewport?: {
        width: number;
        height: number;
    };
    pageUrl?: string;
}
declare function buildReportIssueUrl(env?: ReportIssueEnv): string;
declare function openReportIssue(env?: ReportIssueEnv): void;

export { type ReportIssueEnv, buildReportIssueUrl, openReportIssue };
