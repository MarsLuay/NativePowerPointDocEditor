import { AgentMessage } from '../../agent-types';
export interface AgentChatLogProps {
    messages: AgentMessage[];
    loading?: boolean;
    error?: string | null;
    thinkingLabel?: string;
    workingLabel?: (count: number) => string;
    summaryLabel?: (count: number) => string;
    earlierLabel?: (count: number) => string;
    autoScroll?: boolean;
    humanizeToolName?: (name: string) => string;
    maxVisibleCalls?: number;
    className?: string;
}
declare var __VLS_1: {};
type __VLS_Slots = {} & {
    empty?: (props: typeof __VLS_1) => any;
};
declare const __VLS_component: import('vue').DefineComponent<AgentChatLogProps, {}, {}, {}, {}, import('vue').ComponentOptionsMixin, import('vue').ComponentOptionsMixin, {}, string, import('vue').PublicProps, Readonly<AgentChatLogProps> & Readonly<{}>, {
    className: string;
    error: string | null;
    maxVisibleCalls: number;
    workingLabel: (count: number) => string;
    summaryLabel: (count: number) => string;
    earlierLabel: (count: number) => string;
    loading: boolean;
    thinkingLabel: string;
    autoScroll: boolean;
    humanizeToolName: (name: string) => string;
}, {}, {}, {}, string, import('vue').ComponentProvideOptions, false, {}, any>;
declare const _default: __VLS_WithSlots<typeof __VLS_component, __VLS_Slots>;
export default _default;
type __VLS_WithSlots<T, S> = T & {
    new (): {
        $slots: S;
    };
};
