(function () {
  "use strict";

  var DATA_URL = "data.json";
  var MANIFEST_URL = "site-manifest.json";

  var state = {
    data: null,
    pkg: [],
    sup: [],
    fac: [],
    fromKey: 0,
    toKey: 0
  };

  function byId(id) {
    return document.getElementById(id);
  }

  function fillSelect(el, values, selected) {
    el.innerHTML = "";
    values.forEach(function (v) {
      var opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      opt.selected = selected.indexOf(v) >= 0;
      el.appendChild(opt);
    });
  }

  function fillMonths() {
    var from = byId("from"), to = byId("to");
    from.innerHTML = "";
    to.innerHTML = "";
    state.data.months.forEach(function (period) {
      var label = CartLogic.monthLabel(period);
      ["from", "to"].forEach(function (id) {
        var opt = document.createElement("option");
        opt.value = CartLogic.keyOf(period);
        opt.textContent = label;
        byId(id).appendChild(opt);
      });
    });
    from.value = String(state.data.months.length ? CartLogic.keyOf(state.data.months[0]) : 0);
    to.value = String(state.data.months.length ? CartLogic.keyOf(state.data.months[state.data.months.length - 1]) : 0);
    state.fromKey = Number(from.value);
    state.toKey = Number(to.value);
  }

  function readSelections() {
    state.pkg = Array.prototype.slice.call(byId("pkg").selectedOptions).map(function (o) { return o.value; });
    state.sup = Array.prototype.slice.call(byId("sup").selectedOptions).map(function (o) { return o.value; });
    state.fac = Array.prototype.slice.call(byId("fac").selectedOptions).map(function (o) { return o.value; });
    state.fromKey = Number(byId("from").value);
    state.toKey = Number(byId("to").value);
  }

  function refreshOptions() {
    var opts = CartLogic.options(state.data, { pkg: state.pkg, sup: state.sup, fac: state.fac });
    state.sup = state.sup.filter(function (v) { return opts.suppliers.indexOf(v) >= 0; });
    opts = CartLogic.options(state.data, { pkg: state.pkg, sup: state.sup, fac: state.fac });
    state.fac = state.fac.filter(function (v) { return opts.factories.indexOf(v) >= 0; });
    fillSelect(byId("pkg"), opts.packagingSuppliers, state.pkg);
    fillSelect(byId("sup"), opts.suppliers, state.sup);
    fillSelect(byId("fac"), opts.factories, state.fac);
  }

  function render() {
    var view = CartLogic.buildView(state.data, state);
    var caption = byId("caption");
    var container = byId("table");
    if (!view.rows.length) {
      caption.textContent = "";
      container.innerHTML = "<p>No data for the selected filters.</p>";
      byId("export").disabled = true;
      return;
    }
    var frmLabel = CartLogic.monthLabel(view.months[0]);
    var toLabel = CartLogic.monthLabel(view.months[view.months.length - 1]);
    caption.textContent = view.rows.length.toLocaleString() + " rows \u00b7 " +
      view.months.length + " months \u00b7 " + frmLabel + " to " + toLabel;
    var html = "<table><thead><tr>";
    view.columns.forEach(function (c) {
      html += "<th>" + c + "</th>";
    });
    html += "</tr></thead><tbody>";
    view.rows.forEach(function (r) {
      html += "<tr>";
      r.forEach(function (v, i) {
        var text = i < 4 ? String(v) : Math.round(v).toLocaleString();
        html += "<td>" + text + "</td>";
      });
      html += "</tr>";
    });
    html += "</tbody></table>";
    container.innerHTML = html;
    byId("export").disabled = false;
  }

  function titleOf() {
    var pkg = state.pkg.length ? state.pkg.join(", ") : "All suppliers";
    var fromLabel = CartLogic.monthLabel(CartLogic.periodOf(state.fromKey));
    var toLabel = CartLogic.monthLabel(CartLogic.periodOf(state.toKey));
    return pkg + " \u2014 " + fromLabel + " to " + toLabel;
  }

  function exportExcel() {
    var view = CartLogic.buildView(state.data, state);
    var wb = CartLogic.toWorkbookData(view, titleOf());
    var workbook = new ExcelJS.Workbook();
    var sheet = workbook.addWorksheet("Report");
    var HEADER_BG = { argb: "FFD9E2F3" }, TOTAL_BG = { argb: "FFE2EFDA" };
    var BORDER = {
      top: { style: "thin", color: { argb: "FFBFBFBF" } },
      bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
      left: { style: "thin", color: { argb: "FFBFBFBF" } },
      right: { style: "thin", color: { argb: "FFBFBFBF" } }
    };
    var headerRow = 2, totalRow = wb.rows.length + 1;
    sheet.mergeCells(1, 1, 1, wb.columns.length);
    var titleCell = sheet.getCell(1, 1);
    titleCell.value = wb.title;
    titleCell.font = { bold: true, size: 14 };
    wb.columns.forEach(function (name, c) {
      var headerCell = sheet.getCell(headerRow, c + 1);
      headerCell.value = name;
      headerCell.font = { bold: true };
      headerCell.fill = { type: "pattern", pattern: "solid", fgColor: HEADER_BG };
      headerCell.alignment = { horizontal: "center" };
      headerCell.border = BORDER;
      var totalCell = sheet.getCell(totalRow, c + 1);
      totalCell.value = wb.rows[wb.rows.length - 1][c];
      totalCell.font = { bold: true };
      totalCell.fill = { type: "pattern", pattern: "solid", fgColor: TOTAL_BG };
      totalCell.border = BORDER;
    });
    for (var r = 1; r < wb.rows.length - 1; r++) {
      for (var c = 0; c < wb.columns.length; c++) {
        var cell = sheet.getCell(r + 2, c + 1);
        cell.value = wb.rows[r][c];
        cell.alignment = { horizontal: c >= 4 ? "right" : "left" };
        cell.border = BORDER;
      }
    }
    wb.widths.forEach(function (w, c) {
      sheet.getColumn(c + 1).width = w;
    });
    sheet.views = [{ state: "frozen", ySplit: 2 }];
    workbook.xlsx.writeBuffer().then(function (buffer) {
      var blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = "carton-report.xlsx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }).catch(function (err) {
      byId("export-error").textContent = "Excel export unavailable: " + err.message;
    });
  }

  function init() {
    byId("export").addEventListener("click", function () {
      byId("export-error").textContent = "";
      try {
        if (typeof ExcelJS === "undefined") {
          throw new Error("Excel library could not be loaded (CDN unreachable).");
        }
        exportExcel();
      } catch (err) {
        byId("export-error").textContent = "Excel export unavailable: " + err.message;
      }
    });

    ["pkg", "sup", "fac"].forEach(function (id) {
      byId(id).addEventListener("change", function () {
        readSelections();
        refreshOptions();
        render();
      });
    });
    ["from", "to"].forEach(function (id) {
      byId(id).addEventListener("change", function () {
        readSelections();
        if (state.fromKey > state.toKey) {
          var tmp = state.fromKey;
          state.fromKey = state.toKey;
          state.toKey = tmp;
          byId("from").value = String(state.fromKey);
          byId("to").value = String(state.toKey);
        }
        render();
      });
    });

    Promise.all([
      fetch(DATA_URL).then(function (r) { return r.json(); }),
      fetch(MANIFEST_URL).then(function (r) { return r.json(); })
    ]).then(function (results) {
      state.data = results[0];
      var manifest = results[1];
      byId("footer").textContent = "Data as of " + manifest.published;
      if (!state.data.months.length) {
        byId("caption").textContent = "No data found in the workbook.";
        byId("export").disabled = true;
        return;
      }
      fillMonths();
      refreshOptions();
      render();
    }).catch(function (err) {
      byId("data-error").textContent = "Could not load site data: " + err.message;
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();