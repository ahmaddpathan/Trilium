import {
    getSpreadsheetSheets, parseSpreadsheetWorkbook
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
    }
});
