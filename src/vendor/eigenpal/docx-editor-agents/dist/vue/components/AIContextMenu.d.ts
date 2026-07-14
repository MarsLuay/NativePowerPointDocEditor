export interface AIContextMenuLabels {
    header?: string;
    rewrite?: string;
    expand?: string;
    summarize?: string;
    fixGrammar?: string;
    makeFormal?: string;
    makeCasual?: string;
    translate?: string;
    explain?: string;
    customPlaceholder?: string;
}
export interface AIContextMenuProps {
    isOpen: boolean;
    position: {
        x: number;
        y: number;
    };
    selectedText: string;
    showCustomPrompt?: boolean;
    labels?: AIContextMenuLabels;
}
declare const _default: import('vue').DefineComponent<AIContextMenuProps, {}, {}, {}, {}, import('vue').ComponentOptionsMixin, import('vue').ComponentOptionsMixin, {} & {
    close: () => any;
    action: (action: string, customPrompt?: string | undefined) => any;
}, string, import('vue').PublicProps, Readonly<AIContextMenuProps> & Readonly<{
    onClose?: (() => any) | undefined;
    onAction?: ((action: string, customPrompt?: string | undefined) => any) | undefined;
}>, {
    showCustomPrompt: boolean;
    labels: AIContextMenuLabels;
}, {}, {}, {}, string, import('vue').ComponentProvideOptions, false, {}, any>;
export default _default;
