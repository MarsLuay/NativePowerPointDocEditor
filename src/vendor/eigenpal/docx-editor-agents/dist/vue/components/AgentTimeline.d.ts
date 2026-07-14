import { AgentToolCall } from '../../agent-types';
export interface AgentTimelineProps {
    toolCalls: AgentToolCall[];
    streaming?: boolean;
    maxVisibleCalls?: number;
    humanizeName?: (name: string) => string;
    workingLabel?: (count: number) => string;
    summaryLabel?: (count: number) => string;
    earlierLabel?: (count: number) => string;
}
declare const _default: import('vue').DefineComponent<AgentTimelineProps, {}, {}, {}, {}, import('vue').ComponentOptionsMixin, import('vue').ComponentOptionsMixin, {}, string, import('vue').PublicProps, Readonly<AgentTimelineProps> & Readonly<{}>, {
    streaming: boolean;
    maxVisibleCalls: number;
    humanizeName: (name: string) => string;
    workingLabel: (count: number) => string;
    summaryLabel: (count: number) => string;
    earlierLabel: (count: number) => string;
}, {}, {}, {}, string, import('vue').ComponentProvideOptions, false, {}, any>;
export default _default;
