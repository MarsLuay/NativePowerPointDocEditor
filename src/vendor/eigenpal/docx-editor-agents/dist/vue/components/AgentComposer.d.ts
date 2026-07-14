export interface AgentComposerProps {
    modelValue: string;
    disabled?: boolean;
    placeholder?: string;
    sendLabel?: string;
    className?: string;
}
declare var __VLS_1: {};
type __VLS_Slots = {} & {
    footnote?: (props: typeof __VLS_1) => any;
};
declare const __VLS_component: import('vue').DefineComponent<AgentComposerProps, {}, {}, {}, {}, import('vue').ComponentOptionsMixin, import('vue').ComponentOptionsMixin, {} & {
    submit: () => any;
    "update:modelValue": (next: string) => any;
}, string, import('vue').PublicProps, Readonly<AgentComposerProps> & Readonly<{
    onSubmit?: (() => any) | undefined;
    "onUpdate:modelValue"?: ((next: string) => any) | undefined;
}>, {
    className: string;
    disabled: boolean;
    placeholder: string;
    sendLabel: string;
}, {}, {}, {}, string, import('vue').ComponentProvideOptions, false, {}, any>;
declare const _default: __VLS_WithSlots<typeof __VLS_component, __VLS_Slots>;
export default _default;
type __VLS_WithSlots<T, S> = T & {
    new (): {
        $slots: S;
    };
};
