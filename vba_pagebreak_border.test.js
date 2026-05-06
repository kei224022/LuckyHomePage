/**
 * Test suite for VBA macro "印刷範囲の改ページ直前に黒線を自動設定"
 *
 * Run with:  node vba_pagebreak_border.test.js
 */

const {
    createSheet,
    applyPageBreakBorders,
    parsePrintArea,
    xlContinuous,
    xlNone,
    xlMedium,
    xlThick,
} = require("./vba_pagebreak_border");

// ---- Tiny test runner ----
let passed = 0;
let failed = 0;

function assert(condition, label) {
    if (condition) {
        console.log(`  ✓ ${label}`);
        passed++;
    } else {
        console.error(`  ✗ ${label}`);
        failed++;
    }
}

function describe(name, fn) {
    console.log(`\n[${name}]`);
    fn();
}

// ---- Helpers ----
function borderAt(sheet, row) {
    return sheet.borders[`${row}`] || null;
}

// ============================================================
// Tests
// ============================================================

describe("parsePrintArea", () => {
    const r = parsePrintArea("B2:F20");
    assert(r.startCol === 2,  "startCol = 2 (B)");
    assert(r.startRow === 2,  "startRow = 2");
    assert(r.endCol   === 6,  "endCol   = 6 (F)");
    assert(r.endRow   === 20, "endRow   = 20");

    const r2 = parsePrintArea("A1:Z100");
    assert(r2.startCol === 1,   "A → 1");
    assert(r2.endCol   === 26,  "Z → 26");
});

// ---- TC-01: 印刷範囲が未設定 ----
describe("TC-01: 印刷範囲が未設定の場合エラーメッセージを返す", () => {
    const sheet = createSheet(null, []);
    const result = applyPageBreakBorders(sheet);
    assert(result.ok === false, "ok = false");
    assert(result.message === "印刷範囲が設定されていません。", "エラーメッセージ一致");
});

// ---- TC-02: 改ページなし ----
describe("TC-02: 改ページなし → ボーダー変化なし", () => {
    const sheet = createSheet("A1:E30", []);
    const result = applyPageBreakBorders(sheet);
    assert(result.ok === true, "ok = true");
    assert(Object.keys(sheet.borders).length === 0, "ボーダーは設定されない");
});

// ---- TC-03: 単一改ページ ----
describe("TC-03: 単一改ページ → 改ページ直前行にボーダー設定", () => {
    // 印刷範囲 A1:E30, 改ページが15行目から開始 → 14行目にボーダー
    const sheet = createSheet("A1:E30", [15]);
    const result = applyPageBreakBorders(sheet);
    assert(result.ok === true, "ok = true");
    const b = borderAt(sheet, 14);
    assert(b !== null,                    "14行目にボーダーが設定される");
    assert(b.lineStyle === xlContinuous,  "lineStyle = xlContinuous");
    assert(b.weight    === xlMedium,      "weight = xlMedium");
    assert(b.color.r === 0 && b.color.g === 0 && b.color.b === 0, "color = black");
    assert(borderAt(sheet, 15) === null,  "15行目にはボーダーなし");
});

// ---- TC-04: 複数改ページ ----
describe("TC-04: 複数改ページ → それぞれの直前行にボーダー設定", () => {
    const sheet = createSheet("A1:E50", [11, 21, 31]);
    applyPageBreakBorders(sheet);
    assert(borderAt(sheet, 10) !== null, "10行目にボーダー");
    assert(borderAt(sheet, 20) !== null, "20行目にボーダー");
    assert(borderAt(sheet, 30) !== null, "30行目にボーダー");
    assert(borderAt(sheet, 11) === null, "11行目にはボーダーなし");
});

// ---- TC-05: 既存のMediumボーダーをクリアしてから再設定 ----
describe("TC-05: 既存 Medium/Thick ボーダーを削除してから改ページボーダーを設定", () => {
    // 5行目に既存Mediumボーダー、改ページは10行目から
    const sheet = createSheet("A1:E30", [10], {
        "5":  { lineStyle: xlContinuous, weight: xlMedium },
        "15": { lineStyle: xlContinuous, weight: xlThick  },
    });
    applyPageBreakBorders(sheet);

    const b5  = borderAt(sheet, 5);
    const b15 = borderAt(sheet, 15);
    assert(b5  === null || b5.lineStyle  === xlNone, "5行目のMediumボーダーが削除される");
    assert(b15 === null || b15.lineStyle === xlNone, "15行目のThickボーダーが削除される");

    const b9 = borderAt(sheet, 9);
    assert(b9 !== null && b9.lineStyle === xlContinuous, "9行目(改ページ直前)にボーダー設定");
});

// ---- TC-06: 細線ボーダーは削除しない ----
describe("TC-06: 既存の細線(Thin)ボーダーは削除しない", () => {
    const xlThin = 2;
    const sheet = createSheet("A1:E30", [10], {
        "5": { lineStyle: xlContinuous, weight: xlThin },
    });
    applyPageBreakBorders(sheet);
    const b = borderAt(sheet, 5);
    assert(b !== null && b.weight === xlThin, "細線ボーダーはそのまま保持される");
});

// ---- TC-07: 印刷範囲外の改ページは無視 ----
describe("TC-07: 印刷範囲外の改ページを無視する", () => {
    // 印刷範囲 A5:E25、改ページが3行目(範囲外)と30行目(範囲外)
    const sheet = createSheet("A5:E25", [3, 30]);
    applyPageBreakBorders(sheet);
    assert(borderAt(sheet, 2)  === null, "2行目(印刷範囲外)にボーダーなし");
    assert(borderAt(sheet, 29) === null, "29行目(印刷範囲外)にボーダーなし");
});

// ---- TC-08: 改ページが印刷範囲の最終行+1 → 最終行にボーダー ----
describe("TC-08: 印刷範囲末尾ちょうどの改ページ", () => {
    // 印刷範囲 A1:E20, 改ページが21行目 → 20行目(=endRow)にボーダー
    const sheet = createSheet("A1:E20", [21]);
    applyPageBreakBorders(sheet);
    assert(borderAt(sheet, 20) !== null, "最終行(20行目)にボーダーが設定される");
});

// ---- TC-09: 改ページが印刷範囲の先頭行 → ボーダーは範囲外なので設定しない ----
describe("TC-09: 改ページが印刷範囲の先頭行と同じ場合", () => {
    // 印刷範囲 A5:E20, 改ページが5行目 → borderRow=4(範囲外)
    const sheet = createSheet("A5:E20", [5]);
    applyPageBreakBorders(sheet);
    assert(borderAt(sheet, 4) === null, "4行目(印刷範囲外)にボーダーなし");
});

// ---- TC-10: 複数回実行しても冪等 ----
describe("TC-10: 同じシートで2回実行しても結果が同じ(冪等性)", () => {
    const sheet = createSheet("A1:E30", [15]);
    applyPageBreakBorders(sheet);
    applyPageBreakBorders(sheet);  // 2回目
    const b = borderAt(sheet, 14);
    assert(b !== null && b.lineStyle === xlContinuous && b.weight === xlMedium,
        "2回目実行後も14行目のボーダーが正しい");
    // 他の行にボーダーが増えていないことを確認
    const unexpectedRows = Object.keys(sheet.borders).filter(k => k !== "14" && sheet.borders[k].lineStyle === xlContinuous);
    assert(unexpectedRows.length === 0, "意図しない行にボーダーが追加されていない");
});

// ---- Summary ----
console.log(`\n${"=".repeat(50)}`);
console.log(`結果: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
