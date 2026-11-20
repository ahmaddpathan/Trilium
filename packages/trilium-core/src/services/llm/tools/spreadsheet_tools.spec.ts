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
