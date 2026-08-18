const assert = require("assert");
const Logic = require("../site_template/logic.js");

const DATA = {
  columns: ["SL", "Packaging Supplier", "Supplier", "Factory", "Jan-26", "Feb-26", "Total"],
  months: [[2026, 1], [2026, 2]],
  rows: [
    [1, "M&U", "S1", "F1", 10, 20, 30],
    [2, "M&U", "S2", "F2", 5, 0, 5],
    [3, "Union", "S1", "F1", 7, 7, 14],
    [4, "M&U", "S1", "F3", 0, 0, 0],
  ],
};

assert.deepStrictEqual(Logic.monthLabel([2026, 1]), "Jan-26");
assert.deepStrictEqual(Logic.monthLabel([2026, 12]), "Dec-26");
assert.deepStrictEqual(Logic.keyOf([2026, 1]), 2026 * 12);
assert.deepStrictEqual(Logic.periodOf(2026 * 12 + 1), [2026, 2]);
assert.deepStrictEqual(Logic.rangeMonths([2026, 11], [2027, 1]), [[2026, 11], [2026, 12], [2027, 1]]);
assert.deepStrictEqual(Logic.whole(3.9), 4);

const optsAll = Logic.options(DATA, { pkg: [], sup: [], fac: [] });
assert.deepStrictEqual(optsAll.packagingSuppliers, ["M&U", "Union"]);
assert.deepStrictEqual(optsAll.suppliers, ["S1", "S2"]);
assert.deepStrictEqual(optsAll.factories, ["F1", "F2", "F3"]);

const optsMU = Logic.options(DATA, { pkg: ["M&U"], sup: [], fac: [] });
assert.deepStrictEqual(optsMU.suppliers, ["S1", "S2"]);
assert.deepStrictEqual(optsMU.factories, ["F1", "F2", "F3"]);

const optsMU_S1 = Logic.options(DATA, { pkg: ["M&U"], sup: ["S1"], fac: [] });
assert.deepStrictEqual(optsMU_S1.factories, ["F1", "F3"]);

const view = Logic.buildView(DATA, { pkg: ["M&U"], sup: [], fac: [], fromKey: Logic.keyOf([2026, 2]), toKey: Logic.keyOf([2026, 2]) });
assert.deepStrictEqual(view.columns, ["SL", "Packaging Supplier", "Supplier", "Factory", "Feb-26", "Total"]);
assert.deepStrictEqual(view.rows, [
  [1, "M&U", "S1", "F1", 20, 20],
]);
assert.deepStrictEqual(view.months, [[2026, 2]]);

const viewF3 = Logic.buildView(DATA, { pkg: ["M&U"], sup: ["S1"], fac: ["F3"], fromKey: Logic.keyOf([2026, 1]), toKey: Logic.keyOf([2026, 2]) });
assert.deepStrictEqual(viewF3.rows, []);

const wb = Logic.toWorkbookData(view, "M&U — Feb-26 to Feb-26");
assert.deepStrictEqual(wb.columns, view.columns);
assert.deepStrictEqual(wb.rows[0], view.columns);
assert.deepStrictEqual(wb.rows[1], [1, "M&U", "S1", "F1", 20, 20]);
assert.deepStrictEqual(wb.rows[2], ["", "Total", "", "", 20, 20]);
assert.deepStrictEqual(wb.widths, [4, 20, 10, 9, 8, 7]);

console.log("All logic tests passed");