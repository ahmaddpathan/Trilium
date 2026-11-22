import { describe, expect, it } from "vitest";

import {
    columnIndex,
    columnLetters,
    formatA1Range,
    MAX_RANGE_CELLS,
    parseA1Range,
    rangeCellCount,
    readRange,
    writeRange
} from "./range_address.js";
import type { IWorksheetData } from "./workbook_model.js";

/** The team's "MCP Spreadsheet Test" sheet: a header row, then products and prices (B3 = 1200). */
const sheet: IWorksheetData = {
    id: "sheet-1",
    name: "Sheet1",
    cellData: {
        0: { 0: { v: "A" }, 1: { v: "B" } },
        1: { 0: { v: "Product" }, 1: { v: "Price" } },
        2: { 0: { v: "Laptop" }, 1: { v: 1200, t: 2 } },
        3: { 0: { v: "Mouse" }, 1: { v: 50, t: 2 } },
        4: { 0: { v: "Keyboard" }, 1: { v: 100, t: 2 } },
        5: { 1: { v: 1350, f: "=SUM(B3:B5)", t: 2 } },
        6: { 0: { s: "bold" } }
    }
};

function range(input: string) {
    const result = parseA1Range(input);
    if (!result.ok) throw new Error(result.error);
    return result.range;
}

describe("parseA1Range", () => {
    it("parses a single cell", () => {
        expect(range("B3")).toEqual({ startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 });
    });

    it("parses a rectangle", () => {
        expect(range("A2:B5")).toEqual({ startRow: 1, endRow: 4, startColumn: 0, endColumn: 1 });
    });

    it("accepts lowercase, spaces and $ markers", () => {
        expect(range(" a2:$B$5 ")).toEqual(range("A2:B5"));
    });

    it("normalizes a reversed range", () => {
        expect(range("B5:A2")).toEqual(range("A2:B5"));
    });

    it.each(["", "B", "3", "A0", "A2:", ":B5", "A2:B5:C6", "A:A", "2:2", "A1B2", "R1C1", "AAAA1", "XFE1", "A1048577"])(
        "rejects %j",
        (input) => {
            const result = parseA1Range(input);
            expect(result.ok).toBe(false);
        }
    );

    it("accepts the last cell of the grid", () => {
        expect(range("XFD1048576")).toEqual({ startRow: 1_048_575, endRow: 1_048_575, startColumn: 16_383, endColumn: 16_383 });
    });

    it("rejects a range over the cell limit and says how big it was", () => {
        const result = parseA1Range("A1:Z1000");
        expect(result).toEqual({ ok: false, error: expect.stringContaining("26000 cells") });
    });

    it("allows a range exactly at the limit", () => {
        expect(parseA1Range("A1:J1000").ok).toBe(true);
        expect(rangeCellCount(range("A1:J1000"))).toBe(MAX_RANGE_CELLS);
    });

    it("uses a custom limit when given", () => {
        expect(parseA1Range("A1:B2", 3).ok).toBe(false);
        expect(parseA1Range("A1:B2", 4).ok).toBe(true);
    });
});

describe("column letters", () => {
    it.each([
        [0, "A"],
        [25, "Z"],
        [26, "AA"],
        [701, "ZZ"],
        [702, "AAA"],
        [16_383, "XFD"]
    ])("%i <-> %s", (index, letters) => {
        expect(columnLetters(index)).toBe(letters);
        expect(columnIndex(letters)).toBe(index);
    });
});

describe("formatA1Range", () => {
    it("round-trips what parseA1Range accepts", () => {
        for (const input of ["B3", "A2:B5", "AA10:AB12", "XFD1048576"]) {
            expect(formatA1Range(range(input))).toBe(input);
        }
    });
});

describe("readRange", () => {
    it("reads B3 as 1200", () => {
        expect(readRange(sheet, range("B3"))).toEqual([[{ v: 1200, t: 2 }]]);
    });

    it("returns only the requested rectangle, row by row", () => {
        expect(readRange(sheet, range("A2:B5"))).toEqual([
            [{ v: "Product" }, { v: "Price" }],
            [{ v: "Laptop" }, { v: 1200, t: 2 }],
            [{ v: "Mouse" }, { v: 50, t: 2 }],
            [{ v: "Keyboard" }, { v: 100, t: 2 }]
        ]);
    });

    it("returns the formula alongside its cached value", () => {
        expect(readRange(sheet, range("B6"))).toEqual([[{ v: 1350, f: "=SUM(B3:B5)", t: 2 }]]);
    });

    it("returns null for empty, style-only and out-of-data cells, keeping the grid shape", () => {
        expect(readRange(sheet, range("A6:C7"))).toEqual([
            [null, { v: 1350, f: "=SUM(B3:B5)", t: 2 }, null],
            [null, null, null]
        ]);
    });

    it("returns the plain text of a rich-text cell", () => {
        const linked: IWorksheetData = {
            id: "s",
            name: "s",
            cellData: { 0: { 0: { p: { body: { dataStream: "Trilium\r\n" } } } } }
        };
        expect(readRange(linked, range("A1"))).toEqual([[{ v: "Trilium" }]]);
    });

    it("handles a sheet with no cell data", () => {
        expect(readRange({ id: "s", name: "s", cellData: {} }, range("A1:B2"))).toEqual([
            [null, null],
            [null, null]
        ]);
    });
});

/**
 * A priced list with styles, a merged title, row and column metadata, a rich-text link, a plain
 * formula in B6 and a shared formula: C3 holds it and C4:C5 reuse it through `si`.
 */
function editableSheet(): IWorksheetData {
    return {
        id: "sheet-1",
        name: "Sheet1",
        rowCount: 20,
        columnCount: 5,
        mergeData: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }],
        rowData: { 1: { h: 30 } },
        columnData: { 1: { w: 120 } },
        cellData: {
            0: { 0: { v: "Inventory", s: "title" } },
            1: { 0: { v: "Product", t: 1 }, 1: { v: "Price", t: 1 } },
            2: {
                0: { v: "Laptop", t: 1 },
                1: { v: 1200, t: 2, s: "money" },
                2: { v: 1440, t: 2, f: "=B3*1.2", si: "tax" }
            },
            3: {
                0: { v: "Mouse", t: 1, p: { body: { dataStream: "Mouse\r\n" } } },
                1: { s: { bg: { rgb: "#ffff00" } } },
                2: { v: 0, t: 2, si: "tax" }
            },
            4: { 0: { v: "Keyboard", t: 1 }, 1: { v: 100, t: 2 }, 2: { v: 120, t: 2, si: "tax" } },
            5: { 1: { v: 1300, t: 2, f: "=SUM(B3:B5)", s: "money" } }
        }
    } as IWorksheetData;
}

function written(input: string, values: unknown[][], sheet = editableSheet()) {
    const result = writeRange(sheet, range(input), values);
    if (!result.ok) throw new Error(result.error);
    return result.sheet;
}

function writeError(input: string, values: unknown[][], sheet = editableSheet()) {
    const result = writeRange(sheet, range(input), values);
    return result.ok ? null : result.error;
}

describe("writeRange", () => {
    it("writes only the range, keeping styles, neighbors and sheet metadata", () => {
        const sheet = editableSheet();
        const original = structuredClone(sheet);
        const result = written("A3:B4", [["Tablet", 300], ["Pen", 2.5]], sheet);

        expect(readRange(result, range("A2:B5"))).toEqual([
            [{ v: "Product", t: 1 }, { v: "Price", t: 1 }],
            [{ v: "Tablet", t: 1 }, { v: 300, t: 2 }],
            [{ v: "Pen", t: 1 }, { v: 2.5, t: 2 }],
            [{ v: "Keyboard", t: 1 }, { v: 100, t: 2 }]
        ]);
        expect(result.cellData[2][1]).toEqual({ v: 300, t: 2, s: "money" });
        expect(result.cellData[3][1]).toEqual({ v: 2.5, t: 2, s: { bg: { rgb: "#ffff00" } } });
        expect(result.cellData[2][2]).toBe(sheet.cellData[2][2]);
        expect(result.cellData[3][2]).toBe(sheet.cellData[3][2]);
        expect(result.cellData[0]).toBe(sheet.cellData[0]);
        expect(result.cellData[5]).toBe(sheet.cellData[5]);
        expect({ ...result, cellData: null }).toEqual({ ...original, cellData: null });
        expect(sheet).toEqual(original);
    });

    it("replaces a cell's value, type, rich text and formula, and keeps its other fields", () => {
        const sheet = editableSheet();
        sheet.cellData[3][0] = { ...sheet.cellData[3][0], custom: { tag: "kept" } } as never;
        const result = written("A4:B6", [[true, "x"], [false, 0], [null, 7]], sheet);

        expect(result.cellData[3][0]).toEqual({ v: true, t: 3, custom: { tag: "kept" } });
        expect(result.cellData[5][1]).toEqual({ v: 7, t: 2, s: "money" });
        expect(readRange(result, range("B6"))).toEqual([[{ v: 7, t: 2 }]]);
    });

    it("clears values with null, keeping styled cells and dropping cells left empty", () => {
        const sheet = editableSheet();
        const result = written("A1:B4", [
            [null, null], [null, null], [null, null], [null, null]
        ], sheet);

        expect(result.cellData[0]).toEqual({ 0: { s: "title" } });
        expect(result.cellData[1]).toBeUndefined();
        expect(result.cellData[2]).toEqual({ 1: { s: "money" }, 2: sheet.cellData[2][2] });
        expect(result.cellData[3]).toEqual({
            1: { s: { bg: { rgb: "#ffff00" } } }, 2: sheet.cellData[3][2]
        });
        expect(readRange(result, range("A1:B4")).flat().every((cell) => cell === null)).toBe(true);
        expect(written("E10", [[null]], sheet).cellData).toEqual(sheet.cellData);
        expect(written("E10", [["new"]], sheet).cellData[9]).toEqual({ 4: { v: "new", t: 1 } });
    });

    it("rejects a values grid that does not match the range", () => {
        expect(writeError("A1:B3", [["a", "b"], ["c", "d"]]))
            .toBe('Range "A1:B3" is 3 rows by 2 columns, but values has 2 rows.');
        expect(writeError("A1:B3", [["a", "b"], ["c"], ["e", "f"]]))
            .toBe('Range "A1:B3" is 3 rows by 2 columns, but values row 2 has 1 cell.');
        expect(writeError("B3", [[1, 2]]))
            .toBe('Range "B3" is 1 row by 1 column, but values row 1 has 2 cells.');
        expect(writeError("B3", [])).toBe('Range "B3" is 1 row by 1 column, but values has 0 rows.');
        expect(writeError("B3", "x" as never)).toContain("but values is not an array of rows");
        expect(writeError("B3", ["x"] as never)).toContain("but values row 1 is not an array");
    });

    it("rejects formulas and values other than numbers, strings, booleans and null", () => {
        expect(writeError("A2:B2", [["ok", "=SUM(B3:B5)"]]))
            .toBe('Cell B2 starts with "=": writing formulas is not supported. Write a plain value instead.');
        for (const value of [Number.NaN, Number.POSITIVE_INFINITY, undefined, {}, [1]]) {
            expect(writeError("C7", [[value]]))
                .toBe("Cell C7 has an unsupported value. Use a number, a string, a boolean or null.");
        }
    });

    it("rejects a range over the cell limit or outside the sheet", () => {
        const big = { ...editableSheet(), rowCount: 2000, columnCount: 20 };
        const tooMany = { startRow: 0, endRow: 1000, startColumn: 0, endColumn: 9 };
        expect(writeRange(big, tooMany, [])).toEqual({
            ok: false, error: 'Range "A1:J1001" covers 10010 cells; the limit is 10000. Write a smaller range.'
        });
        expect(writeError("A21", [[1]]))
            .toBe('Range "A21" is outside sheet "Sheet1", which has 20 rows and 5 columns.');
        expect(writeError("F1", [[1]]))
            .toBe('Range "F1" is outside sheet "Sheet1", which has 20 rows and 5 columns.');

        const undeclared: IWorksheetData = { id: "s", name: "Plain", cellData: {} };
        expect(written("T1000", [[1]], undeclared).cellData[999][19]).toEqual({ v: 1, t: 2 });
        expect(writeError("U1", [[1]], undeclared))
            .toBe('Range "U1" is outside sheet "Plain", which has 1000 rows and 20 columns.');
    });

    it("writes the top-left cell of a merge, but no value into the cells it hides", () => {
        expect(written("A1:C1", [["Stock", null, null]]).cellData[0]).toEqual({
            0: { v: "Stock", t: 1, s: "title" }
        });
        expect(writeError("A1:C1", [["Stock", "hidden", null]]))
            .toBe("Cell B1 is hidden by the merged cells A1:C1. Write to A1 instead.");
    });

    it("keeps shared formulas intact, refusing to overwrite the cell that holds one alone", () => {
        expect(writeError("C3", [[0]]))
            .toBe("Cell C3 holds a shared formula that C4 also uses. "
                + "Include every cell that uses it in the range, or leave C3 out.");
        expect(writeError("B3:C4", [[1, 2], [3, 4]])).toContain("that C5 also uses");

        const member = written("C4", [[50]]);
        expect(member.cellData[3][2]).toEqual({ v: 50, t: 2 });
        expect(member.cellData[2][2]).toEqual({ v: 1440, t: 2, f: "=B3*1.2", si: "tax" });
        expect(member.cellData[4][2]).toEqual({ v: 120, t: 2, si: "tax" });

        const group = written("C3:C5", [[1], [2], [3]]);
        expect(readRange(group, range("C3:C5"))).toEqual([
            [{ v: 1, t: 2 }], [{ v: 2, t: 2 }], [{ v: 3, t: 2 }]
        ]);
        expect(JSON.stringify(group.cellData)).not.toContain("tax");
    });
});
