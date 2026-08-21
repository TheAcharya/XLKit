# XLKit Documentation

This folder contains the **structured manual** (chapters under `Manual/`) and this index.

## Manual

- **[Manual index](Manual/README.md)** — Table of contents and links to all chapters (overview, architecture, core model, CSV/images/formatting, security, **CoreUtils** including worksheet protection passwords, testing, examples, full API tables).

**Highlights (1.1.8):** Interned cell formats and streamed worksheet XML (large sheets are much lighter on memory); `setCellFormat` / `removeCellFormat`; **87** unit tests — see chapters [01](Manual/01-Overview-and-Installation.md), [02](Manual/02-Architecture-Modules-and-Source-Map.md), [03](Manual/03-Core-Model-Workbook-Sheet-and-Cells.md), [07](Manual/07-Formatting-Numbers-Alignment-and-Borders.md), [10](Manual/10-Testing-Test-Runner-CI-and-Code-Style.md), and [12](Manual/12-Complete-API-Reference.md).

**Earlier (1.1.6+):** `CoreUtils.configureSheetPassword`, legacy/modern hash helpers, `XLKitTestRunner sheet-password`, and the 11-sheet **`comprehensive`** demo — see chapters [03](Manual/03-Core-Model-Workbook-Sheet-and-Cells.md), [09](Manual/09-Errors-CoreUtils-and-iOS.md), [10](Manual/10-Testing-Test-Runner-CI-and-Code-Style.md), and [12](Manual/12-Complete-API-Reference.md).

## Project links

- **[Main README](../README.md)** — Badges, installation one-liner, quick start.
- **[ARCHITECTURE.md](../ARCHITECTURE.md)** — Module stack, save pipeline, conventions, diagrams.
- **[GUARDRAILS.md](../GUARDRAILS.md)** — Must / must-not constraints for contributors and agents.
- **[AGENT.MD](../AGENT.MD)** — Maintainer and AI-agent development guide (architecture, testing expectations, security).
- **[Tests README](../Tests/README.md)** — Unit test layout (`XLKitTests`, **87** tests).
- **[Test-Workflows README](../Test-Workflows/README.md)** — CLI-generated `.xlsx` outputs; comprehensive demo sheets and unprotect passwords.
- **[Test runner README](../Sources/XLKitTestRunner/README.md)** — Full `XLKitTestRunner` reference (`comprehensive`, `sheet-password`, etc.).
