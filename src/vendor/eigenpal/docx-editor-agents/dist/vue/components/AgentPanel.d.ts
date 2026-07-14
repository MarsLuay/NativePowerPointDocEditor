export interface AgentPanelProps {
    title?: string;
    closeLabel?: string;
    resizeHandleLabel?: string;
    width?: number;
    defaultWidth?: number;
    minWidth?: number;
    maxWidth?: number;
    closed?: boolean;
    /**
     * Show the header close button. Vue parity for React's
     * `onClose=undefined → no button` pattern: pass `:closable="false"` to
     * hide. Defaults to `true` because Vue's `defineEmits` always declares
     * the `close` event regardless of whether the parent attached `@close`.
     */
    closable?: boolean;
    className?: string;
}
declare var __VLS_1: {}, __VLS_3: {};
type __VLS_Slots = {} & {
    icon?: (props: typeof __VLS_1) => any;
} & {
    default?: (props: typeof __VLS_3) => any;
};
declare const __VLS_component: import('vue').DefineComponent<AgentPanelProps, {}, {}, {}, {}, import('vue').ComponentOptionsMixin, import('vue').ComponentOptionsMixin, {} & {
    close: () => any;
    "update:width": (w: number) => any;
}, string, import('vue').PublicProps, Readonly<AgentPanelProps> & Readonly<{
    onClose?: (() => any) | undefined;
    "onUpdate:width"?: ((w: number) => any) | undefined;
}>, {
    title: string;
    closeLabel: string;
    resizeHandleLabel: string;
    width: number;
    defaultWidth: number;
    minWidth: number;
    maxWidth: number;
    closed: boolean;
    closable: boolean;
    className: string;
}, {}, {}, {}, string, import('vue').ComponentProvideOptions, false, {}, any>;
declare const _default: __VLS_WithSlots<typeof __VLS_component, __VLS_Slots>;
export default _default;
type __VLS_WithSlots<T, S> = T & {
    new (): {
        $slots: S;
    };
};
