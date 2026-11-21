/**
 * A1-style range addressing ("B3", "A2:B5") and bounded range reads over the shared workbook model.
 *
 * Used by the MCP spreadsheet tools so an assistant can read a few cells without receiving the whole
 * workbook JSON. Everything here is synchronous and pure: callers load and parse the note themselves.
 */

import { getCellDocumentText, type ICellData, type IRange, type IWorksheetData } from "./workbook_model.js";

/** Largest range a single read may cover. Keeps a response proportional to what was asked for. */
export const MAX_RANGE_CELLS = 10_000;

/** Excel's grid limits (column XFD, row 1,048,576), which Univer also uses. */
export const MAX_ROWS = 1_048_576;
export const MAX_COLUMNS = 16_384;

export type RangeParseResult = { ok: true; range: IRange } | { ok: false; error: string };

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
