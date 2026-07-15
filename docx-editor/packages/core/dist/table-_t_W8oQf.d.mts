import { EditorState, Transaction } from 'prosemirror-state';
import { B as BorderPreset } from './borders-BH_GoYU3.mjs';

declare function insertTable(rows: number, cols: number): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function addRowAbove(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function addRowBelow(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function deleteRow(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function addColumnLeft(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function addColumnRight(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function deleteColumn(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function deleteTable(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function selectTable(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function selectRow(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function selectColumn(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function mergeCells(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function splitCell(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;

declare function setCellBorder(side: 'top' | 'bottom' | 'left' | 'right' | 'all', spec: {
    style: string;
    size?: number;
    color?: {
        rgb: string;
    };
} | null, clearOthers?: boolean): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setTableBorders(preset: BorderPreset, borderSpec?: {
    style: string;
    size: number;
    color: {
        rgb: string;
    };
}): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function removeTableBorders(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
declare function setAllTableBorders(state: EditorState, dispatch?: (tr: Transaction) => void, borderSpec?: {
    style: string;
    size: number;
    color: {
        rgb: string;
    };
}): boolean;
declare function setOutsideTableBorders(state: EditorState, dispatch?: (tr: Transaction) => void, borderSpec?: {
    style: string;
    size: number;
    color: {
        rgb: string;
    };
}): boolean;
declare function setInsideTableBorders(state: EditorState, dispatch?: (tr: Transaction) => void, borderSpec?: {
    style: string;
    size: number;
    color: {
        rgb: string;
    };
}): boolean;
declare function setCellVerticalAlign(align: 'top' | 'center' | 'bottom'): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setCellMargins(margins: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
}): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setCellTextDirection(direction: string | null): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function toggleNoWrap(): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setRowHeight(height: number | null, rule?: 'auto' | 'atLeast' | 'exact'): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function toggleHeaderRow(): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function distributeColumns(): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function autoFitContents(): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setTableProperties(props: {
    width?: number | null;
    widthType?: string | null;
    justification?: 'left' | 'center' | 'right' | null;
}): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function applyTableStyle(styleData: {
    styleId: string;
    tableBorders?: Record<string, unknown>;
    conditionals?: Record<string, unknown>;
    look?: Record<string, boolean>;
}): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setCellFillColor(color: string | null): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setTableBorderColor(color: string): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;
declare function setTableBorderWidth(size: number): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean;

export { setTableBorderWidth as A, setTableBorders as B, setTableProperties as C, splitCell as D, toggleHeaderRow as E, toggleNoWrap as F, addColumnLeft as a, addColumnRight as b, addRowAbove as c, addRowBelow as d, applyTableStyle as e, autoFitContents as f, deleteColumn as g, deleteRow as h, deleteTable as i, distributeColumns as j, insertTable as k, selectRow as l, mergeCells as m, selectTable as n, setAllTableBorders as o, setCellBorder as p, setCellFillColor as q, removeTableBorders as r, selectColumn as s, setCellMargins as t, setCellTextDirection as u, setCellVerticalAlign as v, setInsideTableBorders as w, setOutsideTableBorders as x, setRowHeight as y, setTableBorderColor as z };
