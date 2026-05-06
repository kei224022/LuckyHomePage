/**
 * VBA macro logic reproduced in JavaScript for testing.
 *
 * Original VBA: Sub 印刷範囲の改ページ直前に黒線を自動設定()
 *
 * The macro:
 *  1. Validates that a print area is set.
 *  2. Clears existing medium/thick continuous bottom borders inside the print area.
 *  3. Draws a medium black bottom border on the row just before each horizontal page break.
 */

// ---- Constants matching Excel VBA enumerations ----
const xlContinuous = 1;
const xlNone = -4142;
const xlMedium = -4138;
const xlThick = 4;

// ---- Simulated worksheet model ----

function createSheet(printArea, pageBreakRows, initialBorders = {}) {
    return {
        printArea,          // e.g. "B2:F20"
        pageBreakRows,      // array of row numbers where a page break starts
        borders: { ...initialBorders }, // key: "row,col" -> { lineStyle, weight }
    };
}

function parsePrintArea(printArea) {
    // Parse a simple "A1:B2" style range (column letters + row numbers)
    const match = printArea.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
    if (!match) throw new Error("Invalid print area: " + printArea);
    const colToNum = (s) => s.toUpperCase().split("").reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0);
    return {
        startCol: colToNum(match[1]),
        startRow: parseInt(match[2]),
        endCol: colToNum(match[3]),
        endRow: parseInt(match[4]),
    };
}

// ---- Core logic (mirrors the VBA macro) ----

function applyPageBreakBorders(sheet) {
    if (!sheet.printArea) {
        return { ok: false, message: "印刷範囲が設定されていません。" };
    }

    const { startCol, startRow, endCol, endRow } = parsePrintArea(sheet.printArea);

    // Step 1: Remove existing medium/thick bottom borders in print area
    for (let r = startRow; r <= endRow; r++) {
        const key = `${r}`;
        const border = sheet.borders[key];
        if (border && border.lineStyle === xlContinuous) {
            if (border.weight === xlMedium || border.weight === xlThick) {
                sheet.borders[key] = { lineStyle: xlNone };
            }
        }
    }

    // Step 2: Add medium black bottom border just before each page break
    for (const pbRow of sheet.pageBreakRows) {
        const borderRow = pbRow - 1;
        if (borderRow >= startRow && borderRow <= endRow) {
            sheet.borders[`${borderRow}`] = {
                lineStyle: xlContinuous,
                weight: xlMedium,
                color: { r: 0, g: 0, b: 0 },
            };
        }
    }

    return { ok: true, message: "改ページ黒線を更新しました。" };
}

// ---- Export ----
if (typeof module !== "undefined") {
    module.exports = {
        createSheet, applyPageBreakBorders, parsePrintArea,
        xlContinuous, xlNone, xlMedium, xlThick,
    };
}
