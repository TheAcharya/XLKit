# XLKit — Manual (structured)

This manual is split into **chapters** for navigation and maintenance. It mirrors the style used by [Pipeline Neo’s documentation](https://github.com/TheAcharya/pipeline-neo/tree/main/Documentation): overview first, then architecture, features, security, platform notes, testing, examples, and a **complete public API reference**.

**Start here** — read chapters in order, or jump to what you need.

| Ch | Chapter | Contents |
|----|---------|----------|
| [01](01-Overview-and-Installation.md) | Overview & installation | Performance, compatibility, requirements, SPM, quick start |
| [02](02-Architecture-Modules-and-Source-Map.md) | Architecture & source map | Modules, dependencies, save pipeline, every Swift source file |
| [03](03-Core-Model-Workbook-Sheet-and-Cells.md) | Core model | `Workbook`, `Sheet`, `CellValue`, `Cell`, coordinates |
| [04](04-Bulk-Operations-Ranges-Formulas-and-Merges.md) | Bulk ops & ranges | `setRow`/`setColumn`, formulas, `setRange`, merges |
| [05](05-CSV-and-TSV.md) | CSV & TSV | Import/export, `CSVUtils`, swift-textfile |
| [06](06-Images-Embedding-and-Sizing.md) | Images | Formats, embedding, aspect ratio, EMU sizing |
| [07](07-Formatting-Numbers-Alignment-and-Borders.md) | Formatting | `CellFormat`, numbers, alignment, borders, fonts |
| [08](08-Security-and-Validation.md) | Security | `SecurityManager`, rate limits, quarantine, checksums |
| [09](09-Errors-CoreUtils-and-iOS.md) | Errors, CoreUtils, iOS | `XLKitError`, utilities, sandbox & `safeFileURL` |
| [10](10-Testing-Test-Runner-CI-and-Code-Style.md) | Testing & CI | Unit tests, `XLKitTestRunner`, CI, style |
| [11](11-Examples-Cookbook.md) | Examples cookbook | Five end-to-end examples |
| [12](12-Complete-API-Reference.md) | Complete API reference | Every public type and member |

**Coverage:** The manual is aligned with **all public APIs** in the library targets and documents every Swift source file in `Sources/` (see Chapter 2). The **XLKitTestRunner** executable and **unit tests** are described in Chapter 10.

**Other docs:** [Documentation README](../README.md) · [Project README](../../README.md) · [AGENT.MD](../../AGENT.MD) (contributor/agent guide)
