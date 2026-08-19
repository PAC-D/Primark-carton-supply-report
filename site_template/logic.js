(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.CartLogic = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var LABEL_COUNT = 4;

  function keyOf(period) {
    return period[0] * 12 + (period[1] - 1);
  }

  function periodOf(k) {
    return [Math.floor(k / 12), (k % 12) + 1];
  }

  function monthLabel(period) {
    return MONTHS[period[1] - 1] + "-" + String(period[0]).slice(-2);
  }

  function rangeMonths(frm, to) {
    var out = [];
    for (var k = keyOf(frm); k <= keyOf(to); k++) {
      out.push(periodOf(k));
    }
    return out;
  }

  function whole(v) {
    return Math.round(Number(v));
  }

  function listOrNull(x) {
    return x && x.length ? x : null;
  }

  function options(data, sel) {
    var pkgSet = {}, supSet = {}, facSet = {};
    var rows = data.rows;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var p = r[1], s = r[2], f = r[3];
      pkgSet[p] = true;
      if (!listOrNull(sel.pkg) || sel.pkg.indexOf(p) >= 0) {
        supSet[s] = true;
        if (!listOrNull(sel.sup) || sel.sup.indexOf(s) >= 0) {
          facSet[f] = true;
        }
      }
    }
    return {
      packagingSuppliers: Object.keys(pkgSet).sort(),
      suppliers: Object.keys(supSet).sort(),
      factories: Object.keys(facSet).sort()
    };
  }

  function buildView(data, sel) {
    var pkgSel = listOrNull(sel.pkg);
    var supSel = listOrNull(sel.sup);
    var facSel = listOrNull(sel.fac);
    var baseKey = data.months.length ? keyOf(data.months[0]) : 0;
    var start = sel.fromKey - baseKey;
    var end = sel.toKey - baseKey + 1;
    var months = data.months.slice(start, end);

    var columns = ["SL", "Packaging Supplier", "Supplier", "Factory"];
    for (var i = 0; i < months.length; i++) {
      columns.push(monthLabel(months[i]));
    }
    columns.push("Total");

    var rows = [];
    var all = data.rows;
    for (var i = 0; i < all.length; i++) {
      var r = all[i];
      if (pkgSel && pkgSel.indexOf(r[1]) < 0) continue;
      if (supSel && supSel.indexOf(r[2]) < 0) continue;
      if (facSel && facSel.indexOf(r[3]) < 0) continue;
      var total = 0;
      var values = [];
      for (var j = 0; j < months.length; j++) {
        var v = Number(r[4 + start + j]) || 0;
        values.push(v);
        total += v;
      }
      if (total === 0) continue;
      rows.push([rows.length + 1, r[1], r[2], r[3]].concat(values, [total]));
    }
    return { columns: columns, rows: rows, months: months };
  }

  function toWorkbookData(view, title) {
    var columns = view.columns;
    var rows = view.rows;
    var data = [columns.slice()];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var out = r.slice(0, LABEL_COUNT);
      for (var j = LABEL_COUNT; j < r.length; j++) {
        out.push(whole(r[j]));
      }
      data.push(out);
    }
    var totals = [];
    for (var c = 0; c < columns.length; c++) {
      if (c < LABEL_COUNT) {
        totals.push(columns[c] === "Packaging Supplier" ? "Total" : "");
      } else {
        var sum = 0;
        for (var i = 0; i < rows.length; i++) {
          sum += Number(rows[i][c]) || 0;
        }
        totals.push(whole(sum));
      }
    }
    data.push(totals);

    var widths = columns.map(function (name, c) {
      var longest = name.length;
      for (var i = 0; i < data.length; i++) {
        longest = Math.max(longest, String(data[i][c]).length);
      }
      return Math.min(longest + 2, 40);
    });
    return { title: title, columns: columns, rows: data, widths: widths };
  }

  function selectionSummary(selection) {
    return selection && selection.length ? selection.length + " selected" : "All";
  }

  function pruneSelection(selection, available) {
    var out = [];
    for (var i = 0; i < selection.length; i++) {
      if (available.indexOf(selection[i]) >= 0) {
        out.push(selection[i]);
      }
    }
    return out;
  }

  return {
    keyOf: keyOf,
    periodOf: periodOf,
    monthLabel: monthLabel,
    rangeMonths: rangeMonths,
    whole: whole,
    options: options,
    buildView: buildView,
    toWorkbookData: toWorkbookData,
    selectionSummary: selectionSummary,
    pruneSelection: pruneSelection,
    WORKBOOK_COLORS: {
      headerBg: "FF00205B",
      totalBg: "FFD9E2F3",
      titleColor: "FF00205B",
      headerColor: "FFFFFFFF",
      totalColor: "FF00205B"
    }
  };
});