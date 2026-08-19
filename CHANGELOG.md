# Changelog

All notable changes to this project.

## [Unreleased]

## [1.2.0] — 2026-08-19

### Added

- Dropdown filter widgets with multiple selection and a "Select all" option for Packaging Supplier, Supplier, and Factory (mirrors the app's empty-selection = All filtering)

### Changed

- Static site is now served directly at the repository root (`docs/`) instead of under `/site/`

## [1.1.0] — 2026-08-19

### Added

- PACD theme for the static site: glass navbar with PACD and Primark logos, favicon, Outfit font, navy table headers
- Navy/white Excel export palette shared by the app and the site (header `#00205B`, totals `#D9E2F3`)
- Site publishing now copies the theme CSS and logo/favicon assets

## [1.0.0] — 2026-08-19

### Added

- Filterable carton supply report app (Streamlit): Packaging Supplier, Supplier, Factory and From/To month range filters
- Excel export with styled header row, totals row, and whole numbers
- Static-site publishing ("Publish site" button) with browser-based Excel export, deployable to GitHub Pages
