export interface AIResponsePreviewLabels {
    loading?: string;
    original?: string;
    suggested?: string;
    edit?: string;
    discard?: string;
    accept?: string;
    retry?: string;
    close?: string;
    /** Per-action title overrides (e.g. `{ rewrite: 'Перефразировать' }`). */
    actionTitles?: Partial<Record<string, string>>;
}
export interface AIResponsePreviewProps {
    isVisible: boolean;
    originalText: string;
    responseText: string;
    action: string;
    isLoading: boolean;
    error?: string;
    allowEdit?: boolean;
    showDiff?: boolean;
    showRetry?: boolean;
    labels?: AIResponsePreviewLabels;
}
declare const _default: import('vue').DefineComponent<AIResponsePreviewProps, {}, {}, {}, {}, import('vue').ComponentOptionsMixin, import('vue').ComponentOptionsMixin, {} & {
    accept: (text: string) => any;
    reject: () => any;
    retry: () => any;
}, string, import('vue').PublicProps, Readonly<AIResponsePreviewProps> & Readonly<{
    onAccept?: ((text: string) => any) | undefined;
    onReject?: (() => any) | undefined;
    onRetry?: (() => any) | undefined;
}>, {
    error: string;
    labels: AIResponsePreviewLabels;
    allowEdit: boolean;
    showDiff: boolean;
    showRetry: boolean;
}, {}, {}, {}, string, import('vue').ComponentProvideOptions, false, {}, any>;
export default _default;
