// Re-export the framework-agnostic helpers so existing React imports keep
// working; the canonical implementations live in core.
export {
  buildReportIssueUrl,
  openReportIssue,
  type ReportIssueEnv,
} from '@npde/docx-editor-core/utils/reportIssue';
