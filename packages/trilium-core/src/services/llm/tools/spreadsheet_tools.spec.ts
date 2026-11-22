import { beforeEach, describe, expect, it, vi } from "vitest";

import becca from "../../../becca/becca.js";
import { buildNote } from "../../../test/becca_easy_mocking.js";
import { allToolRegistries } from "./index.js";
import { spreadsheetTools } from "./spreadsheet_tools.js";

const content = JSON.stringify({
    version: 1,
    workbook: {
        sheetOrder: ["second", "first"],
        sheets: {
            first: {
                id: "first", name: "Budget", rowCount: 100, columnCount: 26,
                cellData: { 2: { 1: { v: "private cell value" } } }
            },
            second: { id: "second", name: "Hidden", hidden: 1, cellData: {} }
        },
        styles: { secret: { ff: "private font" } },
        resources: [{ name: "private resource", data: "private payload" }]
    }
});

const tool = Array.from(spreadsheetTools).find(([name]) => name === "get_spreadsheet_info")?.[1];
if (!tool) {
    throw new Error("Spreadsheet metadata tool is not registered");
}

describe("get_spreadsheet_info", () => {
    beforeEach(() => becca.reset());

    it("returns ordered compact metadata, including hidden sheets and missing dimensions", () => {
        const note = buildNote({
            id: "book", title: "Assignment", type: "spreadsheet", content
        });
        note.blobId = "current-version";
        const result = tool.execute({ noteId: "book" });

        expect(result).toEqual({
            noteId: "book", title: "Assignment", blobId: "current-version",
            sheets: [
                { id: "second", name: "Hidden", rowCount: null, columnCount: null, hidden: true },
                { id: "first", name: "Budget", rowCount: 100, columnCount: 26, hidden: false }
            ]
        });
        expect(JSON.stringify(result)).not.toMatch(/cellData|styles|resources|private/);
        expect(note.getContent()).toBe(content);
        expect(result).not.toBeInstanceOf(Promise);
    });

    it("is available in the shared registry as a read-only user-data tool", () => {
        expect(allToolRegistries).toContain(spreadsheetTools);
        expect(tool.mutates).not.toBe(true);
        expect("readsUserData" in tool && tool.readsUserData === false).toBe(false);
        expect(tool.inputSchema.safeParse({ noteId: "" }).success).toBe(false);
    });

    it("returns the same error for missing and excluded notes without reading content", () => {
        const note = buildNote({
            id: "excluded", type: "spreadsheet", content, "#excludeFromLlm": "true"
        });
        const readContent = vi.spyOn(note, "getContent");

        expect(tool.execute({ noteId: "missing" })).toEqual({ error: "Note not found" });
        expect(tool.execute({ noteId: "excluded" })).toEqual({ error: "Note not found" });
        expect(readContent).not.toHaveBeenCalled();
    });

    it("rejects locked notes before accessing their content", () => {
        const note = buildNote({ id: "locked", type: "spreadsheet", content });
        note.isContentAvailable = () => false;
        const readContent = vi.spyOn(note, "getContent");

        expect(tool.execute({ noteId: "locked" })).toEqual({
            error: "Note is protected and cannot be read"
        });
        expect(readContent).not.toHaveBeenCalled();
    });

    it("rejects other note types before reading content", () => {
        const note = buildNote({ id: "text", type: "text", content });
        const readContent = vi.spyOn(note, "getContent");

        expect(tool.execute({ noteId: "text" })).toEqual({ error: "Note is not a spreadsheet" });
        expect(readContent).not.toHaveBeenCalled();
    });

    it("reports malformed workbooks without leaking their contents", () => {
        for (const malformed of ["private malformed JSON", "null", "{}", '{"version":1,"workbook":[]}']) {
            const note = buildNote({ type: "spreadsheet", content: malformed });
            const result = tool.execute({ noteId: note.noteId });
            expect(result).toHaveProperty("error", expect.any(String));
            expect(JSON.stringify(result)).not.toContain("private malformed JSON");
            expect(note.getContent()).toBe(malformed);
        }
    });
});

/** Two sheets; "Products" holds a header row, then products and prices (B3 = 1200). */
const rangeContent = JSON.stringify({
    version: 1,
    workbook: {
        sheetOrder: ["products", "notes"],
        sheets: {
            products: {
                id: "products", name: "Products", rowCount: 1000, columnCount: 20,
                cellData: {
                    0: { 0: { v: "Inventory", s: "title" } },
                    1: { 0: { v: "Product", t: 1 }, 1: { v: "Price", t: 1 } },
                    2: { 0: { v: "Laptop", t: 1 }, 1: { v: 1200, t: 2 } },
                    3: { 0: { v: "Mouse", t: 1 }, 1: { s: "highlight" } },
                    4: { 0: { v: "Keyboard", t: 1 }, 1: { v: 100, t: 2 } },
                    5: {
                        0: { v: "Total", t: 1 },
                        1: { v: 1300, f: "=SUM(B3:B5)", t: 2 },
                        2: { v: true, t: 3 }
                    }
                }
            },
            notes: { id: "notes", name: "Notes", cellData: { 2: { 1: { v: "other sheet" } } } }
        },
        styles: { title: { bl: 1 }, highlight: { bg: { rgb: "#ffff00" } } }
    }
});

describe("get_spreadsheet_range", () => {
    beforeEach(() => becca.reset());

    it("reads a single cell, so B3 returns 1200", () => {
        buildNote({ id: "book", type: "spreadsheet", content: rangeContent });
        const result = readCells({ noteId: "book", sheet: "Products", range: "B3" });

        expect(result).toEqual({
            noteId: "book", sheetId: "products", sheetName: "Products", range: "B3",
            cells: [[{ v: 1200, t: 2 }]]
        });
        expect(result).not.toBeInstanceOf(Promise);
        expect(getSpreadsheetTool("get_spreadsheet_range").mutates).not.toBe(true);
    });

    it("returns only the requested rectangle, with empty cells as null", () => {
        buildNote({ id: "book", type: "spreadsheet", content: rangeContent });
        const expected = {
            noteId: "book", sheetId: "products", sheetName: "Products", range: "A2:B5",
            cells: [
                [{ v: "Product", t: 1 }, { v: "Price", t: 1 }],
                [{ v: "Laptop", t: 1 }, { v: 1200, t: 2 }],
                [{ v: "Mouse", t: 1 }, null],
                [{ v: "Keyboard", t: 1 }, { v: 100, t: 2 }]
            ]
        };
        const result = readCells({ noteId: "book", sheet: "Products", range: "A2:B5" });

        expect(result).toEqual(expected);
        expect(readCells({ noteId: "book", sheet: "Products", range: "b5:a2" })).toEqual(expected);
        expect(JSON.stringify(result))
            .not.toMatch(/Inventory|Total|other sheet|cellData|styles|highlight/);
    });

    it("returns text, numbers, booleans and formulas with their cached values", () => {
        buildNote({ id: "book", type: "spreadsheet", content: rangeContent });

        expect(readCells({ noteId: "book", sheet: "Products", range: "A6:D6" })).toMatchObject({
            range: "A6:D6",
            cells: [[
                { v: "Total", t: 1 },
                { v: 1300, f: "=SUM(B3:B5)", t: 2 },
                { v: true, t: 3 },
                null
            ]]
        });
    });

    it("selects a sheet by ID or by name, and reports an unknown sheet", () => {
        buildNote({ id: "book", type: "spreadsheet", content: rangeContent });

        for (const sheet of ["notes", "Notes"]) {
            expect(readCells({ noteId: "book", sheet, range: "B3" })).toEqual({
                noteId: "book", sheetId: "notes", sheetName: "Notes", range: "B3",
                cells: [[{ v: "other sheet" }]]
            });
        }
        expect(readCells({ noteId: "book", sheet: "Missing", range: "B3" }))
            .toEqual({ error: "Spreadsheet sheet not found" });
    });

    it("bounds the response by the range and rejects malformed or over-limit ranges", () => {
        const note = buildNote({ id: "book", type: "spreadsheet", content: rangeContent });
        const { cells } = readCells({ noteId: "book", sheet: "Products", range: "A1:J1000" }) as {
            cells: unknown[][];
        };
        expect(cells).toHaveLength(1000);
        expect(cells.every((row) => row.length === 10)).toBe(true);

        const readContent = vi.spyOn(note, "getContent");
        for (const range of ["", "B", "A0", "A1:B2:C3", "A:A", "R1C1"]) {
            expect(readCells({ noteId: "book", sheet: "Products", range }))
                .toEqual({ error: expect.stringContaining("Invalid range") });
        }
        expect(readCells({ noteId: "book", sheet: "Products", range: "A1:Z1000" }))
            .toEqual({ error: expect.stringMatching(/26000 cells; the limit is 10000/) });
        expect(readContent).not.toHaveBeenCalled();
    });

    it("rejects other note types, excluded, locked and binary notes", () => {
        const notes = [
            buildNote({ id: "text", type: "text", content: rangeContent }),
            buildNote({
                id: "excluded", type: "spreadsheet", content: rangeContent,
                "#excludeFromLlm": "true"
            }),
            buildNote({ id: "locked", type: "spreadsheet", content: rangeContent })
        ];
        notes[2].isContentAvailable = () => false;
        const reads = notes.map((note) => vi.spyOn(note, "getContent"));
        const binary = buildNote({ id: "binary", type: "spreadsheet" });
        binary.getContent = () => new Uint8Array([1, 2, 3]);
        const args = { sheet: "Products", range: "B3" };

        expect(readCells({ ...args, noteId: "text" }))
            .toEqual({ error: "Note is not a spreadsheet" });
        expect(readCells({ ...args, noteId: "excluded" })).toEqual({ error: "Note not found" });
        expect(readCells({ ...args, noteId: "missing" })).toEqual({ error: "Note not found" });
        expect(readCells({ ...args, noteId: "locked" }))
            .toEqual({ error: "Note is protected and cannot be read" });
        expect(readCells({ ...args, noteId: "binary" }))
            .toEqual({ error: "Spreadsheet content must be JSON text" });
        for (const read of reads) {
            expect(read).not.toHaveBeenCalled();
        }
    });
});

describe("set_spreadsheet_range", () => {
    beforeEach(() => becca.reset());

    it("writes only the range, keeps styles, and reads back with the new blobId", () => {
        const note = buildWritableSheet({ id: "book" });
        const result = writeCells({
            noteId: "book", sheet: "Products", range: "a3:b4",
            values: [["Tablet", 300], ["Pen", 2]], expectedBlobId: "blob-0"
        });

        expect(result).toEqual({
            noteId: "book", sheetId: "products", sheetName: "Products", range: "A3:B4",
            blobId: "blob-1"
        });
        expect(result).not.toBeInstanceOf(Promise);
        expect(getSpreadsheetTool("set_spreadsheet_range").mutates).toBe(true);
        expect(note.saveRevision).toHaveBeenCalledTimes(1);
        expect(readCells({ noteId: "book", sheet: "Products", range: "A2:C6" })).toMatchObject({
            cells: [
                [{ v: "Product", t: 1 }, { v: "Price", t: 1 }, null],
                [{ v: "Tablet", t: 1 }, { v: 300, t: 2 }, null],
                [{ v: "Pen", t: 1 }, { v: 2, t: 2 }, null],
                [{ v: "Keyboard", t: 1 }, { v: 100, t: 2 }, null],
                [{ v: "Total", t: 1 }, { v: 1300, f: "=SUM(B3:B5)", t: 2 }, { v: true, t: 3 }]
            ]
        });
        expect(readCells({ noteId: "book", sheet: "notes", range: "B3" }))
            .toMatchObject({ cells: [[{ v: "other sheet" }]] });

        const stored = JSON.parse(String(note.getContent()));
        const original = JSON.parse(rangeContent);
        expect(stored.workbook.sheets.products.cellData[3][1])
            .toEqual({ v: 2, t: 2, s: "highlight" });
        expect(stored.workbook.sheets.notes).toEqual(original.workbook.sheets.notes);
        expect(stored.workbook.styles).toEqual(original.workbook.styles);
        expect(stored.workbook.sheetOrder).toEqual(original.workbook.sheetOrder);
    });

    it("clears cells with null and replaces formulas, chaining writes through blobId", () => {
        const note = buildWritableSheet({ id: "book" });
        const first = writeCells({
            noteId: "book", sheet: "Products", range: "B6:C6",
            values: [[1500, null]], expectedBlobId: "blob-0"
        }) as { blobId: string };
        const second = writeCells({
            noteId: "book", sheet: "Products", range: "A1",
            values: [[null]], expectedBlobId: first.blobId
        });

        expect(second).toMatchObject({ range: "A1", blobId: "blob-2" });
        const read = readCells({ noteId: "book", sheet: "Products", range: "A1:C6" }) as {
            cells: unknown[][];
        };
        expect(read.cells[0]).toEqual([null, null, null]);
        expect(read.cells[5]).toEqual([{ v: "Total", t: 1 }, { v: 1500, t: 2 }, null]);
        const cells = JSON.parse(String(note.getContent())).workbook.sheets.products.cellData;
        expect(cells[0][0]).toEqual({ s: "title" });
        expect(cells[5][1]).toEqual({ v: 1500, t: 2 });
    });

    it("refuses a stale expectedBlobId before reading the note", () => {
        const note = buildWritableSheet({ id: "book" });
        const readContent = vi.spyOn(note, "getContent");

        expect(writeCells({
            noteId: "book", sheet: "Products", range: "B3",
            values: [[1]], expectedBlobId: "blob-from-an-earlier-read"
        })).toEqual({ error: expect.stringMatching(/^Stale expectedBlobId/) });
        expect(readContent).not.toHaveBeenCalled();
        expect(note.setContent).not.toHaveBeenCalled();
        expect(note.saveRevision).not.toHaveBeenCalled();
    });

    it("rejects bad ranges, values and sheets without changing the note", () => {
        const note = buildWritableSheet({ id: "book" });
        const write = ({ sheet = "Products", range, values }: {
            sheet?: string; range: string; values: unknown[][];
        }) => writeCells({ noteId: "book", sheet, range, values, expectedBlobId: "blob-0" });

        for (const range of ["", "A0", "A:A", "R1C1"]) {
            expect(write({ range, values: [[1]] }))
                .toEqual({ error: expect.stringContaining("Invalid range") });
        }
        expect(write({ range: "A1:Z1000", values: [] }))
            .toEqual({ error: expect.stringMatching(/26000 cells; the limit is 10000/) });
        expect(write({ range: "A1:B2", values: [[1, 2]] })).toEqual({
            error: 'Range "A1:B2" is 2 rows by 2 columns, but values has 1 row.'
        });
        expect(write({ range: "B3", values: [["=B4*2"]] }))
            .toEqual({ error: expect.stringContaining("writing formulas is not supported") });
        expect(write({ range: "B1001", values: [[1]] }))
            .toEqual({ error: expect.stringContaining("which has 1000 rows and 20 columns") });
        expect(write({ sheet: "Missing", range: "B3", values: [[1]] }))
            .toEqual({ error: "Spreadsheet sheet not found" });
        expect(note.getContent()).toBe(rangeContent);
        expect(note.setContent).not.toHaveBeenCalled();
        expect(note.saveRevision).not.toHaveBeenCalled();
    });

    it("rejects other note types, excluded, locked and binary notes", () => {
        const notes = [
            buildWritableSheet({ id: "text", type: "text" }),
            buildWritableSheet({ id: "excluded", "#excludeFromLlm": "true" }),
            buildWritableSheet({ id: "locked" })
        ];
        notes[2].isContentAvailable = () => false;
        const binary = buildWritableSheet({ id: "binary" });
        binary.getContent = () => new Uint8Array([1, 2, 3]);
        const args = { sheet: "Products", range: "B3", values: [[1]], expectedBlobId: "blob-0" };

        expect(writeCells({ ...args, noteId: "text" }))
            .toEqual({ error: "Note is not a spreadsheet" });
        expect(writeCells({ ...args, noteId: "excluded" })).toEqual({ error: "Note not found" });
        expect(writeCells({ ...args, noteId: "missing" })).toEqual({ error: "Note not found" });
        expect(writeCells({ ...args, noteId: "locked" }))
            .toEqual({ error: "Note is protected and cannot be modified" });
        expect(writeCells({ ...args, noteId: "binary" }))
            .toEqual({ error: "Spreadsheet content must be JSON text" });
        for (const note of [...notes, binary]) {
            expect(note.setContent).not.toHaveBeenCalled();
        }
    });
});

/**
 * A spreadsheet note holding `rangeContent` in memory: `setContent` stores the new content and
 * moves `blobId` from "blob-0" to "blob-1" and on, as a real save does. `saveRevision` is a stub.
 */
function buildWritableSheet(definition: Parameters<typeof buildNote>[0]) {
    const note = buildNote({ type: "spreadsheet", content: rangeContent, ...definition });
    let content = note.getContent();
    let saves = 0;
    note.blobId = "blob-0";
    note.getContent = () => content;
    note.setContent = vi.fn((next: string | Uint8Array) => {
        content = next;
        note.blobId = `blob-${++saves}`;
    }) as typeof note.setContent;
    note.saveRevision = vi.fn() as typeof note.saveRevision;
    return note;
}

function readCells(args: { noteId: string; sheet: string; range: string }) {
    return getSpreadsheetTool("get_spreadsheet_range").execute(args);
}

function writeCells(args: {
    noteId: string; sheet: string; range: string; values: unknown[][]; expectedBlobId: string;
}) {
    return getSpreadsheetTool("set_spreadsheet_range").execute(args);
}

function getSpreadsheetTool(toolName: string) {
    const found = Array.from(spreadsheetTools).find(([name]) => name === toolName)?.[1];
    if (!found) {
        throw new Error(`Spreadsheet tool ${toolName} is not registered`);
    }
    return found;
}
