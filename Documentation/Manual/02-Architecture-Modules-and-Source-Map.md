# Chapter 02 — Architecture, Modules, and Source Map

Navigation: [← Chapter 01](01-Overview-and-Installation.md) · [Manual index](README.md) · [Chapter 03 →](03-Core-Model-Workbook-Sheet-and-Cells.md)

XLKit is a **Swift Package** with a small main target that re-exports implementation modules. This keeps compile times reasonable and lets advanced users depend on a subset (for example, only `XLKitCore`).

## Module dependency graph

```
XLKit (umbrella, re-exports all)
├── XLKitCore      — Workbook, Sheet, types, CoreUtils, SecurityManager
├── XLKitFormatters — CSVUtils (swift-textfile)
├── XLKitImages    — ImageUtils, ImageSizingUtils
└── XLKitXLSX      — XLSXEngine (ZIPFoundation)

XLKitFormatters → XLKitCore, TextFile (swift-textfile)
XLKitImages → XLKitCore
XLKitXLSX → XLKitCore, XLKitFormatters, XLKitImages, ZIPFoundation
```

**Executable:** `XLKitTestRunner` depends on `XLKit` and **CoreXLSX** (validation only; not linked into the library).

## What each module contains

| Module | Responsibility |
|--------|----------------|
| **XLKitCore** | Domain model: `Workbook`, `Sheet`, `CellValue`, `Cell`, `CellFormat`, `SheetState`, `SheetProtection`, coordinates, `ExcelImage`, `XLKitError`, `CoreUtils`, `SecurityManager`. |
| **XLKitFormatters** | `CSVUtils` — RFC-style CSV and TSV import/export via [swift-textfile](https://github.com/orchetect/swift-textfile). |
| **XLKitImages** | `ImageUtils` (format detection, dimensions, `ExcelImage` construction), `ImageSizingUtils` (aspect fit, EMUs, Excel column/row width formulas). |
| **XLKitXLSX** | `XLSXEngine` — builds the Open XML package (styles, shared strings, worksheets, drawings, relationships) and zips it with ZIPFoundation. |
| **XLKit** | `@_exported import` of the four modules above; `Workbook+API` and `Sheet+API` convenience extensions (save, CSV, image embed). |

The `XLKit` **struct** in `XLKit.swift` is an empty namespace; all behaviour lives on `Workbook` and `Sheet` or static helpers.

## Save pipeline (end-to-end)

1. You build a `Workbook` in memory (sheets, cells, formats, images).
2. `workbook.save(to: url)` (sync or async) calls `XLSXEngine.generateXLSX(workbook:to:)`.
3. The engine applies **SecurityManager** (rate limit, optional path rules), writes OOXML parts under a temporary directory, then creates a **ZIP** archive as `.xlsx`.
4. Workbook XML includes per-sheet `state` attributes (`.hidden` / `.veryHidden`) and `activeTab` on `<workbookView>` when the first sheet is hidden. Worksheet XML includes `<sheetProtection>` after `</sheetData>` when `Sheet.protection` is set.
5. Optional **SHA-256** checksum sidecar if `SecurityManager.enableChecksumStorage` is `true`.

For low-level or test use, you may call `XLSXEngine.generateXLSX` directly; app code should normally use `save(to:)`. `XLSXEngine.formatToKey(_:)` is public for hashing/deduplicating `CellFormat` instances (styles XML uses the same keying).

## Source file map (100% of Swift sources)

| Path | Role |
|------|------|
| `Package.swift` | Targets, platforms, SPM dependencies. |
| `Sources/XLKit/XLKit.swift` | Re-exports. |
| `Sources/XLKit/Workbook+API.swift` | CSV/TSV convenience inits, `save`, import/export. |
| `Sources/XLKit/Sheet+API.swift` | `setRow`/`setColumn`, CSV export, `embedImage*` APIs. |
| `Sources/XLKitCore/CoreTypes.swift` | Core types and `CoreUtils`. |
| `Sources/XLKitCore/SecurityManager.swift` | Rate limiting, logging, quarantine, checksum hooks. |
| `Sources/XLKitFormatters/CSVUtils.swift` | CSV/TSV. |
| `Sources/XLKitImages/ImageUtils.swift` | Image detection and `ExcelImage` creation. |
| `Sources/XLKitImages/ImageSizingUtils.swift` | Sizing and EMU math. |
| `Sources/XLKitXLSX/XLSXEngine.swift` | Full XLSX generation. |
| `Sources/XLKitTestRunner/main.swift` | CLI entry. |
| `Sources/XLKitTestRunner/ExcelGenerators.swift` | Demo generators + CoreXLSX checks. |
| `Sources/XLKitTestRunner/ImageEmbedGenerators.swift` | Image embedding scenarios. |
| `Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift` | Template for new CLI tests. |
| `Tests/XLKitTests/*.swift` | Unit tests (see [Chapter 10](10-Testing-Test-Runner-CI-and-Code-Style.md)). |

Non-Swift artefacts: `SECURITY.md`, `CHANGELOG.md`, `Assets/`, `Test-Data/`, CI workflows under `.github/workflows/`, and Markdown under `Documentation/`.

## Related chapters

- Data model: [Chapter 03 — Core model](03-Core-Model-Workbook-Sheet-and-Cells.md)
- Export details and public engine API: [Chapter 12 — API reference](12-Complete-API-Reference.md) (`XLSXEngine`, `formatToKey`)
