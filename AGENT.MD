# AGENT.MD

## Project: XLKit

**XLKit** is a pure Swift library for creating and writing Microsoft Excel (.xlsx) files on macOS.  
There is no UI—this is a backend/data library only.

---

## Agent Guidelines

### 1. Scope of Work
- Focus on robust, type-safe, and **extremely easy-to-use** Swift code.
- All code must be compatible with Swift 6.0 and macOS 13+.
- No UI code; only library logic, tests, and documentation.
- Support for advanced Excel features (images, GIFs, formulas, merging, column sizing, etc.) is encouraged, but must be implemented in a modular, maintainable, and user-friendly way.

### 2. Coding Standards
- Use modern Swift best practices (async/await, error handling, value types where possible).
- All public APIs must be documented with doc comments and designed for discoverability and ease of use.
- All new features must include corresponding unit tests and usage examples.
- Avoid global state and side effects.
- Prefer fluent, chainable, and bulk APIs where possible.

### 3. Testing
- All code must be covered by unit tests in `Tests/XLKitTests`.
- Tests should be clear, deterministic, and not rely on external resources.
- All new features (including image and column sizing) must be tested.

### 4. Documentation
- Update `README.md` with new features or API changes.
- All public types and methods must have doc comments and usage examples.

### 5. CI/CD
- Ensure the GitHub Actions workflow (`.github/workflows/build.yml`) passes before merging.
- No platform support except macOS.

### 6. File/Folder Structure
- Source code: `Sources/XLKit/`
- Tests: `Tests/XLKitTests/`
- CI: `.github/workflows/`
- No UI, assets, or platform-specific code outside macOS.

---

## Out of Scope

- No UI or app code.
- No Windows or Linux support.
- No dependencies except Swift standard library and Foundation.

---

## Agent Behavior

- Always prefer clarity, safety, and **ease of use** over brevity.
- If unsure, ask for clarification in a comment or PR.
- Do not auto-merge breaking changes.

## Code Style & Formatting Policy

- All code must follow the style rules in `.swift-format`.
- Use 4-space indentation, trailing commas, grouped imports, and modern Swift idioms.
- No force-unwraps or force-casts in public API.
- All public APIs must have doc comments.
- Code must be formatted before PRs are merged.
- Run `swift-format format -i .` to auto-format the codebase.
- See README and `.cursorrules` for more details.

## Modular Structure

- **XLKitCore**: Core types (Workbook, Sheet, Cell, CellValue, CellFormat, etc.)
- **XLKitFormatters**: CSV/TSV import/export utilities (CSVUtils)
- **XLKitImages**: Image utilities (ImageUtils, ExcelImage, ImageFormat)
- **XLKitXLSX**: XLSX/ZIP/XML engine
- **XLKit**: Main API, re-exporting all modules and providing static convenience methods

## CSV/TSV Import/Export
- Use `XLKit.createWorkbookFromCSV`, `XLKit.createWorkbookFromTSV`, `XLKit.exportSheetToCSV`, `XLKit.exportSheetToTSV`, `XLKit.importCSVIntoSheet`, and `XLKit.importTSVIntoSheet` for all CSV/TSV operations.
- These methods are powered by the `XLKitFormatters` module, but are exposed on `XLKit` for ease of use.

## Image Utilities
- Use `ImageUtils` for image format detection, size, and Excel image creation.
- All image-related types (ExcelImage, ImageFormat) are defined in `XLKitCore` and `XLKitImages`.

## Dependency Rules
- `XLKitCore` has no dependencies on other modules.
- `XLKitImages` and `XLKitFormatters` depend on `XLKitCore`.
- `XLKitXLSX` depends on all lower modules.
- `XLKit` re-exports all modules and provides the main API. 