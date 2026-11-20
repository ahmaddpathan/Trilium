import {
    type IWorkbookData,
    type IWorksheetData,
    parseWorkbookData,
    type PersistedData
} from "./workbook_model.js";

/** Validates spreadsheet data for structured access without changing the lenient render parser. */
export function parseSpreadsheetWorkbook(content: string): PersistedData {
    const parsed = parseWorkbookData(content);
    if (!parsed.ok) {
        throw new Error("Spreadsheet content is not valid JSON");
    }
    const data: unknown = parsed.data;
    if (!isRecord(data) || !Number.isSafeInteger(data.version) || Number(data.version) < 0) {
        throw new Error("Invalid spreadsheet version");
    }
    if (!isRecord(data.workbook)) {
        throw new Error("Invalid spreadsheet workbook");
    }
    const workbook = data.workbook;
    if (!isRecord(workbook.sheets)) {
        throw new Error("Invalid spreadsheet sheets");
    }
    if (workbook.name !== undefined && typeof workbook.name !== "string") {
        throw new Error("Invalid spreadsheet workbook name");
    }
    const sheets = Object.fromEntries(Object.entries(workbook.sheets).map(([id, value]) => {
        if (!isRecord(value) || id.length === 0 || value.id !== id
            || typeof value.name !== "string") {
            throw new Error(`Invalid spreadsheet sheet: ${id}`);
        }
        for (const dimension of ["rowCount", "columnCount"]) {
            const count = value[dimension];
            if (count !== undefined && (!Number.isSafeInteger(count) || Number(count) < 0)) {
                throw new Error(`Invalid ${dimension} for spreadsheet sheet: ${id}`);
            }
        }
        if (value.hidden !== undefined && value.hidden !== 0 && value.hidden !== 1) {
            throw new Error(`Invalid hidden flag for spreadsheet sheet: ${id}`);
        }
        if (value.cellData !== undefined && !isRecord(value.cellData)) {
            throw new Error(`Invalid cell data for spreadsheet sheet: ${id}`);
        }
        // Empty historical sheets can omit cellData; preserve every other stored field.
        return [id, { ...value, cellData: value.cellData ?? {} }] as const;
    }));
    const result = { ...data, workbook: { ...workbook, sheets } } as PersistedData;
    orderedSheetIds(result.workbook);
    return result;
}

/** Lists all sheets, including hidden ones, without returning any cell content. */
export function getSpreadsheetSheets(workbook: IWorkbookData) {
    return orderedSheetIds(workbook).map((id) => {
        const sheet = workbook.sheets[id];
        return {
            id,
            name: sheet.name,
            rowCount: sheet.rowCount ?? null,
            columnCount: sheet.columnCount ?? null,
            hidden: Boolean(sheet.hidden)
        };
    });
}

/** Resolves an exact stored ID first, or an unambiguous exact sheet name. */
export function findSpreadsheetSheet(workbook: IWorkbookData, selector: string): IWorksheetData {
    if (Object.hasOwn(workbook.sheets, selector)) {
        return workbook.sheets[selector];
    }
    const matches = Object.values(workbook.sheets).filter((sheet) => sheet.name === selector);
    if (matches.length === 0) {
        throw new Error("Spreadsheet sheet not found");
    }
    if (matches.length > 1) {
        throw new Error("Spreadsheet sheet name is ambiguous; use its ID");
    }
    return matches[0];
}

function orderedSheetIds(workbook: IWorkbookData): string[] {
    const ids = Object.keys(workbook.sheets);
    if (workbook.sheetOrder === undefined) {
        return ids;
    }
    if (!Array.isArray(workbook.sheetOrder)
        || workbook.sheetOrder.some((id) => typeof id !== "string"
            || !Object.hasOwn(workbook.sheets, id))
        || new Set(workbook.sheetOrder).size !== workbook.sheetOrder.length) {
        throw new Error("Invalid spreadsheet sheet order");
    }
    const ordered = new Set(workbook.sheetOrder);
    return [...workbook.sheetOrder, ...ids.filter((id) => !ordered.has(id))];
}

function isRecord(value: unknown): value is Record<string, unknown> {
    return value !== null && typeof value === "object" && !Array.isArray(value);
}
