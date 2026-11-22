/**
 * A1-style range addressing ("B3", "A2:B5") and bounded range reads and writes over the shared
 * workbook model.
 *
 * Used by the MCP spreadsheet tools so an assistant can read or write a few cells without handling
 * the whole workbook JSON. Everything here is synchronous and pure: callers load, parse and save
 * the note themselves.
 */

import {
    type CellMatrix,
    CellValueType,
    getCellDocumentText,
    type ICellData,
    type IRange,
    type IWorksheetData
} from "./workbook_model.js";

/** Largest range a single read or write may cover, so a call stays proportional to its range. */
export const MAX_RANGE_CELLS = 10_000;

/** Excel's grid limits (column XFD, row 1,048,576), which Univer also uses. */
export const MAX_ROWS = 1_048_576;
export const MAX_COLUMNS = 16_384;

/** Univer's grid size for a sheet that declares no `rowCount` or `columnCount`. */
const DEFAULT_ROW_COUNT = 1000;
const DEFAULT_COLUMN_COUNT = 20;

/** The cell fields a write replaces. Every other field, such as the style `s`, is kept. */
const CELL_CONTENT_KEYS = ["v", "t", "p", "f", "si", "ft", "fd", "ref", "xf"] as const;

export type RangeParseResult = { ok: true; range: IRange } | { ok: false; error: string };

/** A value `writeRange` stores in a cell. `null` clears the cell's value. */
export type RangeWriteValue = string | number | boolean | null;

export type RangeWriteResult = { ok: true; sheet: IWorksheetData } | { ok: false; error: string };

/** One cell in a range read. `null` in the result grid means the cell holds nothing. */
export interface RangeCell {
    /** The stored value. For rich-text cells (links, mixed formatting) this is the plain text. */
    v: string | number | boolean | null;
    /** The formula, when the cell has one (e.g. "=SUM(A1:A3)"). `v` is then its last cached result. */
    f?: string;
    /** Univer's value type (1 string, 2 number, 3 boolean, 4 forced string), when stored. */
    t?: number;
}

const CELL_PATTERN = /^\$?([A-Z]{1,3})\$?([0-9]{1,7})$/;

/**
 * Parses "B3" or "A2:B5" (case-insensitive, `$` markers allowed) into a 0-based inclusive range.
 * A reversed range such as "B5:A2" is normalized, as in Excel. Whole-column or whole-row ranges
 * ("A:A", "2:2") are not supported, so a read is always a bounded rectangle.
 */
export function parseA1Range(input: string, maxCells = MAX_RANGE_CELLS): RangeParseResult {
    const parts = input.trim().toUpperCase().split(":");
    if (parts.length > 2) {
        return { ok: false, error: `Invalid range "${input}". Use a cell like "B3" or a range like "A2:B5".` };
    }

    const cells = parts.map(parseCell);
    if (cells.some((cell) => cell === null)) {
        return { ok: false, error: `Invalid range "${input}". Use a cell like "B3" or a range like "A2:B5".` };
    }

    const [start, end = start] = cells as { row: number; column: number }[];
    const range: IRange = {
        startRow: Math.min(start.row, end.row),
        endRow: Math.max(start.row, end.row),
        startColumn: Math.min(start.column, end.column),
        endColumn: Math.max(start.column, end.column)
    };

    const count = rangeCellCount(range);
    if (count > maxCells) {
        return { ok: false, error: `Range "${input}" covers ${count} cells; the limit is ${maxCells}. Request a smaller range.` };
    }
    return { ok: true, range };
}

/** Formats a 0-based inclusive range back to A1 ("B3" for one cell, else "A2:B5"). */
export function formatA1Range(range: IRange): string {
    const start = `${columnLetters(range.startColumn)}${range.startRow + 1}`;
    if (range.startRow === range.endRow && range.startColumn === range.endColumn) return start;
    return `${start}:${columnLetters(range.endColumn)}${range.endRow + 1}`;
}

/** 0-based column index to its letters (0 -> "A", 25 -> "Z", 26 -> "AA"). */
export function columnLetters(index: number): string {
    let n = index;
    let letters = "";
    do {
        letters = String.fromCharCode(65 + (n % 26)) + letters;
        n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return letters;
}

/** Column letters to their 0-based index ("A" -> 0, "Z" -> 25, "AA" -> 26). Letters must be uppercase. */
export function columnIndex(letters: string): number {
    let n = 0;
    for (const char of letters) {
        n = n * 26 + (char.charCodeAt(0) - 64);
    }
    return n - 1;
}

export function rangeCellCount(range: IRange): number {
    return (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1);
}

/**
 * Reads a rectangle of cells from a sheet, row by row. The grid always has the range's exact shape,
 * so `rows[i][j]` is the cell at (startRow + i, startColumn + j). Cells that hold nothing (missing,
 * or only a style) are `null`. Cells outside the sheet's stored data read as `null`, not as an error.
 */
export function readRange(sheet: IWorksheetData, range: IRange): (RangeCell | null)[][] {
    const rows: (RangeCell | null)[][] = [];
    for (let row = range.startRow; row <= range.endRow; row++) {
        const cells: (RangeCell | null)[] = [];
        for (let column = range.startColumn; column <= range.endColumn; column++) {
            cells.push(toRangeCell(sheet.cellData?.[row]?.[column]));
        }
        rows.push(cells);
    }
    return rows;
}

/**
 * Writes a grid of values into a rectangle of a sheet and returns the updated sheet, leaving the
 * input untouched. `values[i][j]` goes to the cell at (startRow + i, startColumn + j), so the grid
 * must have the range's exact shape. A written cell loses its old value, rich text and formula and
 * keeps its style and other fields; `null` clears the value and drops a cell left with nothing.
 * Cells outside the range, merges and row and column metadata stay as they are.
 *
 * Returns an error instead when the range exceeds the cell limit or the sheet's grid, a value is a
 * formula or not a plain value, a value would land in a cell hidden by a merge, or the range covers
 * the cell holding a shared formula but not every cell that uses it, since those cells lose their
 * formula once it is gone.
 */
export function writeRange(
    sheet: IWorksheetData,
    range: IRange,
    values: unknown[][]
): RangeWriteResult {
    const address = formatA1Range(range);
    const count = rangeCellCount(range);
    if (count > MAX_RANGE_CELLS) {
        const error = `Range "${address}" covers ${count} cells; the limit is ${MAX_RANGE_CELLS}.`;
        return { ok: false, error: `${error} Write a smaller range.` };
    }
    const rowCount = sheet.rowCount ?? DEFAULT_ROW_COUNT;
    const columnCount = sheet.columnCount ?? DEFAULT_COLUMN_COUNT;
    if (range.endRow >= rowCount || range.endColumn >= columnCount) {
        return {
            ok: false,
            error: `Range "${address}" is outside sheet "${sheet.name}", which has `
                + `${rowCount} rows and ${columnCount} columns.`
        };
    }

    const error = getShapeError(range, values, address)
        ?? getValueError(range, values)
        ?? getMergeError(sheet, range, values)
        ?? getSharedFormulaError(sheet, range);
    if (error) {
        return { ok: false, error };
    }

    const cellData: CellMatrix = { ...sheet.cellData };
    for (const [i, rowValues] of values.entries()) {
        const row = range.startRow + i;
        const cells = { ...cellData[row] };
        for (const [j, value] of rowValues.entries()) {
            const column = range.startColumn + j;
            const cell = toWrittenCell(cells[column], value as RangeWriteValue);
            if (cell) {
                cells[column] = cell;
            } else {
                delete cells[column];
            }
        }
        if (Object.keys(cells).length > 0) {
            cellData[row] = cells;
        } else {
            delete cellData[row];
        }
    }
    return { ok: true, sheet: { ...sheet, cellData } };
}

function parseCell(text: string): { row: number; column: number } | null {
    const match = CELL_PATTERN.exec(text);
    if (!match) return null;

    const column = columnIndex(match[1]);
    const row = Number(match[2]) - 1;
    if (row < 0 || row >= MAX_ROWS || column >= MAX_COLUMNS) return null;
    return { row, column };
}

function toRangeCell(cell: ICellData | undefined): RangeCell | null {
    if (!cell) return null;

    const text = cell.v == null || cell.v === "" ? getCellDocumentText(cell) : "";
    const v = cell.v != null && cell.v !== "" ? cell.v : text || null;
    if (v === null && !cell.f) return null;

    const result: RangeCell = { v };
    if (cell.f) result.f = cell.f;
    if (cell.t != null) result.t = cell.t;
    return result;
}

function getShapeError(range: IRange, values: unknown[][], address: string): string | null {
    const rows = range.endRow - range.startRow + 1;
    const columns = range.endColumn - range.startColumn + 1;
    const expected = `Range "${address}" is ${plural(rows, "row")} by `
        + `${plural(columns, "column")}, but values`;
    if (!Array.isArray(values)) return `${expected} is not an array of rows.`;
    if (values.length !== rows) return `${expected} has ${plural(values.length, "row")}.`;

    for (const [i, row] of values.entries()) {
        if (!Array.isArray(row)) return `${expected} row ${i + 1} is not an array.`;
        if (row.length !== columns) {
            return `${expected} row ${i + 1} has ${plural(row.length, "cell")}.`;
        }
    }
    return null;
}

function getValueError(range: IRange, values: unknown[][]): string | null {
    for (const [i, row] of values.entries()) {
        for (const [j, value] of row.entries()) {
            const cell = cellName(range.startRow + i, range.startColumn + j);
            if (typeof value === "string" && value.startsWith("=")) {
                return `Cell ${cell} starts with "=": writing formulas is not supported. `
                    + "Write a plain value instead.";
            }
            const supported = value === null || typeof value === "string"
                || typeof value === "boolean" || Number.isFinite(value);
            if (!supported) {
                return `Cell ${cell} has an unsupported value. `
                    + "Use a number, a string, a boolean or null.";
            }
        }
    }
    return null;
}

/** Univer shows only the top-left cell of a merge, so a value in its other cells is hidden. */
function getMergeError(sheet: IWorksheetData, range: IRange, values: unknown[][]): string | null {
    for (const merge of Array.isArray(sheet.mergeData) ? sheet.mergeData : []) {
        const startRow = Math.max(merge.startRow, range.startRow);
        const endRow = Math.min(merge.endRow, range.endRow);
        const startColumn = Math.max(merge.startColumn, range.startColumn);
        const endColumn = Math.min(merge.endColumn, range.endColumn);
        for (let row = startRow; row <= endRow; row++) {
            for (let column = startColumn; column <= endColumn; column++) {
                const isTopLeft = row === merge.startRow && column === merge.startColumn;
                const value = values[row - range.startRow][column - range.startColumn];
                if (!isTopLeft && value !== null) {
                    const topLeft = cellName(merge.startRow, merge.startColumn);
                    return `Cell ${cellName(row, column)} is hidden by the merged cells `
                        + `${formatA1Range(merge)}. Write to ${topLeft} instead.`;
                }
            }
        }
    }
    return null;
}

/**
 * Univer drops the formula of every cell that shares one through `si` when the cell holding it is
 * gone (`buildSheetFormulaData` in `@univerjs/engine-formula`). Rebinding the group to another cell
 * means rewriting the formula's references, which needs Univer's formula lexer.
 */
function getSharedFormulaError(sheet: IWorksheetData, range: IRange): string | null {
    const holders = new Map<string, string>();
    for (let row = range.startRow; row <= range.endRow; row++) {
        for (let column = range.startColumn; column <= range.endColumn; column++) {
            const cell = sheet.cellData[row]?.[column];
            if (cell?.si && typeof cell.f === "string" && cell.f.startsWith("=")) {
                holders.set(cell.si, cellName(row, column));
            }
        }
    }
    if (holders.size === 0) return null;

    for (const [rowKey, cells] of Object.entries(sheet.cellData)) {
        for (const [columnKey, cell] of Object.entries(cells ?? {})) {
            const holder = cell?.si ? holders.get(cell.si) : undefined;
            const row = Number(rowKey);
            const column = Number(columnKey);
            const inRange = row >= range.startRow && row <= range.endRow
                && column >= range.startColumn && column <= range.endColumn;
            if (holder && !inRange) {
                const user = cellName(row, column);
                return `Cell ${holder} holds a shared formula that ${user} also uses. `
                    + `Include every cell that uses it in the range, or leave ${holder} out.`;
            }
        }
    }
    return null;
}

function toWrittenCell(cell: ICellData | undefined, value: RangeWriteValue): ICellData | undefined {
    const kept: Record<string, unknown> = { ...cell };
    for (const key of CELL_CONTENT_KEYS) {
        delete kept[key];
    }
    if (value !== null) {
        kept.v = value;
        kept.t = typeof value === "number" ? CellValueType.NUMBER
            : typeof value === "boolean" ? CellValueType.BOOLEAN
                : CellValueType.STRING;
    }
    return Object.keys(kept).length > 0 ? kept as ICellData : undefined;
}

function cellName(row: number, column: number): string {
    return `${columnLetters(column)}${row + 1}`;
}

function plural(count: number, word: string): string {
    return `${count} ${word}${count === 1 ? "" : "s"}`;
}
