import { describe, expect, it } from "vitest";

import {
    findSpreadsheetSheet,
    getSpreadsheetSheets,
    parseSpreadsheetWorkbook
} from "./workbook_access.js";

function fixture() {
    return {
        version: 1,
        pluginState: { enabled: true },
        workbook: {
            name: "Budget",
            sheetOrder: ["second", "first"],
            sheets: {
                first: {
                    id: "first", name: "Expenses", rowCount: 1000, columnCount: 20,
                    cellData: { 2: { 1: { v: 1200, f: "=SUM(B1:B2)" } } }
                },
                second: {
                    id: "second", name: "Hidden", hidden: 1, cellData: {},
                    customState: { keep: true }
                }
            },
            styles: { bold: { bl: 1 } },
            resources: [{ name: "plugin", data: "preserve this" }]
        }
    };
}

describe("spreadsheet workbook access", () => {
    it("lists ordered metadata, includes hidden sheets, and omits all cell and plugin content", () => {
        const source = fixture();
        const original = JSON.stringify(source);
        const parsed = parseSpreadsheetWorkbook(original);

        expect(getSpreadsheetSheets(parsed.workbook)).toEqual([
            { id: "second", name: "Hidden", rowCount: null, columnCount: null, hidden: true },
            { id: "first", name: "Expenses", rowCount: 1000, columnCount: 20, hidden: false }
        ]);
        expect(parsed).toEqual(source);
        expect(JSON.stringify(source)).toBe(original);
        expect(findSpreadsheetSheet(parsed.workbook, "first").cellData[2][1].v).toBe(1200);
        expect(findSpreadsheetSheet(parsed.workbook, "Expenses"))
            .toBe(parsed.workbook.sheets.first);
    });

    it("uses exact IDs before names, and rejects ambiguous or missing names", () => {
        const source = fixture();
        source.workbook.sheets.second.name = "first";
        const parsed = parseSpreadsheetWorkbook(JSON.stringify(source));
        expect(findSpreadsheetSheet(parsed.workbook, "first").name).toBe("Expenses");
        expect(() => findSpreadsheetSheet(parsed.workbook, "Hidden")).toThrow("not found");
        expect(() => findSpreadsheetSheet(parsed.workbook, "expenses")).toThrow("not found");
        expect(() => findSpreadsheetSheet(parsed.workbook, " Expenses ")).toThrow("not found");
        expect(() => findSpreadsheetSheet(parsed.workbook, "")).toThrow("not found");

        source.workbook.sheets.second.name = "Expenses";
        const ambiguous = parseSpreadsheetWorkbook(JSON.stringify(source));
        expect(() => findSpreadsheetSheet(ambiguous.workbook, "Expenses"))
            .toThrow("ambiguous");
        expect(findSpreadsheetSheet(ambiguous.workbook, "second").id).toBe("second");
    });

    it("uses stored IDs safely even when they match prototype property names", () => {
        const parsed = parseSpreadsheetWorkbook(JSON.stringify({
            version: 1,
            workbook: {
                sheets: Object.fromEntries(["__proto__", "toString"].map((id) => [
                    id, { id, name: `Sheet ${id}`, cellData: {} }
                ]))
            }
        }));
        expect(getSpreadsheetSheets(parsed.workbook).map((sheet) => sheet.id))
            .toEqual(["__proto__", "toString"]);
        expect(findSpreadsheetSheet(parsed.workbook, "__proto__").name).toBe("Sheet __proto__");
        expect(findSpreadsheetSheet(parsed.workbook, "toString").name).toBe("Sheet toString");
        expect(() => findSpreadsheetSheet(parsed.workbook, "constructor")).toThrow("not found");
    });

    it("supports historical order and empty sheets without dropping unlisted sheets", () => {
        const source = fixture();
        const { sheetOrder, ...workbook } = source.workbook;
        const historical = parseSpreadsheetWorkbook(JSON.stringify({ version: 1, workbook }));
        expect(getSpreadsheetSheets(historical.workbook).map((sheet) => sheet.id))
            .toEqual(["first", "second"]);

        const partial = parseSpreadsheetWorkbook(JSON.stringify({
            version: 1, workbook: { ...workbook, sheetOrder: [sheetOrder[0]] }
        }));
        expect(getSpreadsheetSheets(partial.workbook).map((sheet) => sheet.id))
            .toEqual(["second", "first"]);

        const empty = parseSpreadsheetWorkbook(JSON.stringify({
            version: 1,
            workbook: { sheetOrder: ["blank"], sheets: { blank: { id: "blank", name: "Blank" } } }
        }));
        expect(findSpreadsheetSheet(empty.workbook, "blank").cellData).toEqual({});
        expect(getSpreadsheetSheets(empty.workbook)[0])
            .toMatchObject({ rowCount: null, columnCount: null });
        expect(getSpreadsheetSheets(parseSpreadsheetWorkbook(JSON.stringify({
            version: 1, workbook: { sheetOrder: [], sheets: {} }
        })).workbook)).toEqual([]);
    });

    it("rejects malformed JSON and malformed top-level or sheet structures", () => {
        expect(() => parseSpreadsheetWorkbook("{broken")).toThrow("valid JSON");
        const invalid = [
            null, [], "text", 12, {}, { version: "1" }, { version: -1 }, { version: 1.5 },
            { version: 1, workbook: null },
            { version: 1, workbook: [] },
            { version: 1, workbook: { sheets: [] } },
            { version: 1, workbook: { sheets: null } },
            { version: 1, workbook: { name: 12, sheets: {} } }
        ];
        for (const value of invalid) {
            expect(() => parseSpreadsheetWorkbook(JSON.stringify(value))).toThrow();
        }

        for (const sheet of [null, [], "text", {}, { id: "other", name: "Wrong" },
            { id: "sheet", name: 123 }]) {
            expect(() => parseSpreadsheetWorkbook(JSON.stringify({
                version: 1, workbook: { sheets: { sheet } }
            }))).toThrow("sheet");
        }
    });

    it("rejects invalid dimensions, flags, cell containers, and sheet order", () => {
        for (const dimension of ["rowCount", "columnCount"]) {
            for (const count of [-1, 1.5, "20", null, Number.MAX_SAFE_INTEGER + 1]) {
                expect(() => parseSpreadsheetWorkbook(JSON.stringify({
                    version: 1,
                    workbook: { sheets: { s: { id: "s", name: "S", [dimension]: count } } }
                }))).toThrow(dimension);
            }
        }
        for (const hidden of [null, "1", 2, false]) {
            expect(() => parseSpreadsheetWorkbook(JSON.stringify({
                version: 1, workbook: { sheets: { s: { id: "s", name: "S", hidden } } }
            }))).toThrow("hidden");
        }
        for (const cellData of [null, [], "text", 0]) {
            expect(() => parseSpreadsheetWorkbook(JSON.stringify({
                version: 1, workbook: { sheets: { s: { id: "s", name: "S", cellData } } }
            }))).toThrow("cell data");
        }
        for (const sheetOrder of [null, "first", ["first", "first"], ["missing"], [1]]) {
            expect(() => parseSpreadsheetWorkbook(JSON.stringify({
                ...fixture(), workbook: { ...fixture().workbook, sheetOrder }
            }))).toThrow("sheet order");
        }
    });
});
