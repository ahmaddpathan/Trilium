import { describe, expect, it } from "vitest";

import {
    columnIndex,
    columnLetters,
    formatA1Range,
    MAX_RANGE_CELLS,
    parseA1Range,
    rangeCellCount,
    readRange
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
