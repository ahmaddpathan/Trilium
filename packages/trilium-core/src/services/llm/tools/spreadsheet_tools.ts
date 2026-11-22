import {
    formatA1Range, MAX_RANGE_CELLS, parseA1Range, readRange
} from "@triliumnext/commons/src/lib/spreadsheet/range_address.js";
import {
    findSpreadsheetSheet, getSpreadsheetSheets, parseSpreadsheetWorkbook
} from "@triliumnext/commons/src/lib/spreadsheet/workbook_access.js";
import { z } from "zod";

import { getLlmNote } from "./helpers.js";
import { defineTools } from "./tool_registry.js";

export const spreadsheetTools = defineTools({
    get_spreadsheet_info: {
        description: [
            "Get a spreadsheet note's sheet IDs, names, declared dimensions and visibility.",
            "Use this to choose a sheet before reading or updating a range.",
            "Returns compact metadata and the current blobId, without cell contents.",
            "A null rowCount or columnCount means the workbook does not declare that dimension."
        ].join(" "),
        inputSchema: z.object({
            noteId: z.string().min(1).describe("The ID of the spreadsheet note")
        }),
        execute: ({ noteId }: { noteId: string }) => {
            const note = getLlmNote(noteId);
            if (!note) {
                return { error: "Note not found" };
            }
            if (!note.isContentAvailable()) {
                return { error: "Note is protected and cannot be read" };
            }
            if (note.type !== "spreadsheet") {
                return { error: "Note is not a spreadsheet" };
            }
            const content = note.getContent();
            if (typeof content !== "string") {
                return { error: "Spreadsheet content must be JSON text" };
            }

            try {
                const data = parseSpreadsheetWorkbook(content);
                return {
                    noteId: note.noteId,
                    title: note.getTitleOrProtected(),
                    blobId: note.blobId,
                    sheets: getSpreadsheetSheets(data.workbook)
                };
            } catch (error) {
                return {
                    error: error instanceof Error ? error.message : "Invalid spreadsheet workbook"
                };
            }
        }
    },

    get_spreadsheet_range: {
        description: [
            "Read a rectangle of cells from one sheet of a spreadsheet note, in A1 notation",
            `such as 'B3' or 'A2:B5'. A read covers at most ${MAX_RANGE_CELLS} cells;`,
            "whole columns or rows ('A:A', '2:2') are not supported.",
            "Use get_spreadsheet_info first to find the sheet's ID or name.",
            "cells holds the range row by row: cells[i][j] is the cell i rows below and",
            "j columns right of the range's top-left cell. An empty cell is null.",
            "Otherwise v is the stored value, f the formula when there is one",
            "(v is then its last computed result), and t the stored value type",
            "(1 string, 2 number, 3 boolean, 4 forced string)."
        ].join(" "),
        inputSchema: z.object({
            noteId: z.string().min(1).describe("The ID of the spreadsheet note"),
            sheet: z.string().min(1)
                .describe("The sheet's ID, or its exact name if no other sheet shares it"),
            range: z.string().describe("The cells to read in A1 notation, e.g. 'B3' or 'A2:B5'")
        }),
        execute: ({ noteId, sheet, range }: { noteId: string; sheet: string; range: string }) => {
            const note = getLlmNote(noteId);
            if (!note) {
                return { error: "Note not found" };
            }
            if (!note.isContentAvailable()) {
                return { error: "Note is protected and cannot be read" };
            }
            if (note.type !== "spreadsheet") {
                return { error: "Note is not a spreadsheet" };
            }
            const parsedRange = parseA1Range(range);
            if (!parsedRange.ok) {
                return { error: parsedRange.error };
            }
            const content = note.getContent();
            if (typeof content !== "string") {
                return { error: "Spreadsheet content must be JSON text" };
            }

            try {
                const { workbook } = parseSpreadsheetWorkbook(content);
                const worksheet = findSpreadsheetSheet(workbook, sheet);
                return {
                    noteId: note.noteId,
                    sheetId: worksheet.id,
                    sheetName: worksheet.name,
                    range: formatA1Range(parsedRange.range),
                    cells: readRange(worksheet, parsedRange.range)
                };
            } catch (error) {
                return {
                    error: error instanceof Error ? error.message : "Invalid spreadsheet workbook"
                };
            }
        }
    }
});
