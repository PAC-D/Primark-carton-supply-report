# PACD Restyle Follow-up: Dropdown Filters with Multi-Select and Select All

Date: 2026-08-19
Status: Approved design (user: yes to model, approach A, widget spec)

## Context

The static site's Packaging Supplier / Supplier / Factory filters are native
`<select multiple size="6">` list boxes (`site_template/index.html:31,35,39`),
which are tall and require Ctrl/Cmd-click to multi-select. The user wants:

1. All filters rendered as dropdowns.
2. Packaging Supplier, Supplier and Factory to support multiple selection
   with a "Select all" option.

The From/To filters are already dropdowns (`custom-select`, index.html:43,47)
and are unchanged by this work.

Scope is the **static site only** (`site_template/` + regenerated `docs/site/`).
The Streamlit app (`app.py`) keeps its current `st.multiselect` filters.

## Goal

Match the Streamlit app's filtering model (app.py:74-87) in a dropdown
widget: **empty selection = "All" = no filter**, with chained options
(Supplier options depend on the Packaging Supplier selection; Factory
options depend on both).

## Filter Model (Streamlit Mirror)

- `pkg`, `sup`, `fac` selections are each an array of values; the empty
  array means "All" (no filter). This is already how `CartLogic.buildView`
  and `CartLogic.options` treat the state — no data-layer change.
- Options are chained exactly as today via `CartLogic.options`:
  - `opts.packagingSuppliers` — always full list
  - `opts.suppliers` — narrowed by the `pkg` selection
  - `opts.factories` — narrowed by the `pkg` + `sup` selections
- After a chain change (user touches an upstream filter), downstream
  selections are pruned to still-valid options (today's behavior in
  `refreshOptions`, app.js:58-65). If pruning empties a downstream
  selection, that widget reverts to "All".
- "Select all" is a display affordance over the empty-selection model:
  - Select-all checkbox **checked** ⟺ selection is empty. It always
    matches the current option list (nothing to prune by definition).
  - Picking any specific option unchecks select-all and becomes the
    selection.
  - Unchecking the last selected option re-checks select-all (selection
    back to empty = All).

## Widget Design

### Markup (index.html)

Replace each of the three list boxes with a widget container; From/To
selects are untouched:

```html
<div class="filter-widget" id="pkg-widget" data-option-key="pkg">
  <label class="select-label">Packaging Supplier</label>
  <button type="button" class="filter-trigger" aria-haspopup="listbox" aria-expanded="false">
    <span class="filter-summary">All</span>
    <span class="filter-chevron" aria-hidden="true"></span>
  </button>
  <div class="filter-panel" role="listbox" hidden>
    <label class="filter-all">
      <input type="checkbox" checked> Select all
    </label>
    <div class="filter-options"></div>
  </div>
</div>
```

(`sup-widget`, `fac-widget` identical with their own `data-option-key`.)

### Behavior (app.js — new `makeFilterWidget` factory)

DOM-only widget; all model rules live in `logic.js` pure functions (see
Logic Split below). The widget exposes:

- `setOptions(values)` — rebuild the option rows (used by
  `refreshOptions`), preserving the selection except pruning already done
  by the caller
- `getSelection()` — current array (empty = All)
- `setSelection(values)` — set programmatically, updates checkbox states
  and trigger summary

Widget mechanics:

- Clicking the trigger toggles the panel (hidden ↔ visible) and flips
  `aria-expanded`.
- Clicking outside the widget closes the panel. Esc closes it.
- A checkbox row click toggles that value in the selection:
  - value added → select-all unchecks
  - last value removed → select-all re-checks
- The select-all row toggles between empty (checked, summary "All") and
  all options selected individually (unchecked, summary "N selected").
- Trigger summary text: `"All"` when the selection is empty, otherwise
  `"N selected"`.
- Multiple widgets: opening one closes the others.

### Wiring changes (app.js)

- `fillSelect`/`selectedOptions` reads for pkg/sup/fac (app.js:20,51-53)
  are replaced by widget reads in `readSelections`.
- `refreshOptions` (app.js:58-65) rebuilds only the three widget option
  lists instead of the three selects; the prune logic is unchanged in
  behavior but moves to the `CartLogic.pruneSelection` helper.
- The From/To change handler (app.js:181-193) is untouched.

### Logic Split (logic.js — pure, Node-testable)

New `CartLogic` functions:

- `selectionSummary(selection)` → `"All"` if empty, else `"N selected"`.
- `pruneSelection(selection, available)` → selection filtered to values
  still in `available` (order preserved).

No changes to `buildView`, `options`, `toWorkbookData`, or
`WORKBOOK_COLORS`.

### Styling (style.css)

- `.filter-widget` — position relative; inherits the grid placement the
  selects had (`custom-select` sizing).
- `.filter-trigger` — styled like `.custom-select` (full width, PACD
  border/hover), with a CSS chevron.
- `.filter-panel` — absolute, full-width, white background, PACD shadow,
  max-height ~260px with `overflow-y: auto`, z-index above table.
- `.filter-all` — bold row, navy text, bottom divider.
- `.filter-option` rows — checkbox + label, hover highlight (light navy),
  checked rows keep the PACD accent.
- `.filter-summary` — ellipsis on overflow.

## Error Handling

- No new error surface: widget rows are generated from
  `CartLogic.options` output only, so a row can never reference an
  unavailable option.
- If data.json fails to load, the existing error box shows and widgets
  never initialize (same as today).

## Testing

- `tests/test_logic.js` (node): new assertions for `selectionSummary`
  (empty → "All", 2 values → "2 selected") and `pruneSelection` (valid
  kept, invalid dropped, empty preserved).
- `python -m pytest -q`: existing 39 tests must stay green (no Python
  changes expected).
- `node --check` on `app.js` and `logic.js`.
- Manual serve check (`python -m http.server 8010 --directory docs`, open
  `http://localhost:8010/site/`): dropdown opens/closes, select-all
  toggles, chained narrowing prunes correctly, table updates, From/To
  still swap when inverted.

## Publish Flow

`publish.py` needs no changes (it copies `site_template/` wholesale).
Regenerate `docs/site/` from real data (`Sales Record V1.xlsx`, 424
rows) so the preview matches, byte-identity check for the 7 template
files as before, then commit ("Updated at ..." + feature commits).

## Out of Scope

- Streamlit app filters (explicit user decision)
- Search-as-you-type inside dropdowns
- From/To filter behavior or style
- Excel export changes
- Mobile touch interactions beyond the existing responsive CSS
