# XLKit — Guardrails

Hard constraints for contributors and AI agents. Prefer this file when deciding **what not to do**; prefer [ARCHITECTURE.md](ARCHITECTURE.md) for **how the system is shaped**.

**See also:** [ARCHITECTURE.md](ARCHITECTURE.md), [.cursorrules](.cursorrules), [AGENT.MD](AGENT.MD), [Tests/README.md](Tests/README.md), [SECURITY.md](SECURITY.md).

**Current suite (keep in sync):** **80** Swift Testing tests in `XLKitTests` (15 `@Suite`s + `XLKitTestSupport`); CLI demos via `XLKitTestRunner`.

---

## How to use this document

| Audience | Expectation |
|----------|-------------|
| **Human contributors** | Treat §1–§8 as merge blockers unless a maintainer explicitly waives a rule. |
| **AI agents** | Read this file before structural, XLSX, or image-sizing changes. Do not rationalize around a guardrail; ask or stop. |
| **Both** | When a rule is learned from a regression, add a new entry under §9 (Signs) with Trigger / Instruction / Reason / Provenance. |

**Relationship to other docs**

- **ARCHITECTURE.md** — modules, save pipeline, diagrams, conventions.
- **AGENT.MD / .cursorrules** — living agent briefing (must stay in sync with each other).
- **GUARDRAILS.md** — short, enforceable “never / always” list. Keep it scannable; link out for depth.

---

## 1. Naming & product identity

| Rule | Detail |
|------|--------|
| **XLKit only** | Use XLKit naming in code, comments, symbols, CLI, and logs. No alternate fork / product identifiers in the codebase. |
| **File headers** | New Swift files use the project header (ARCHITECTURE §5.2): XLKit URL line + MIT — no `Created by` lines. |
| **Test suites** | Feature suites are `*Tests.swift` with `@Suite`; shared helpers live in `XLKitTestSupport` (`XLKitTestBase.swift`). |

---

## 2. Module boundaries (non-negotiable)

Extend **bottom-up**. Do not invent OOXML or image math only in the TestRunner.

```text
XLKitCore → Formatters / Images → XLKitXLSX → XLKit (umbrella)
```

| Always | Never |
|--------|-------|
| Put new model / CoreUtils / security types in **XLKitCore** | Duplicate domain types inside Formatters, Images, or XLSX |
| Put CSV/TSV behaviour in **XLKitFormatters** | Reimplement CSV quoting outside swift-textfile |
| Put sizing / EMU / format detection in **XLKitImages** | Hardcode column width, row height, or EMUs in XLSXEngine or TestRunner |
| Put package XML / ZIP in **XLKitXLSX** | Emit ad hoc OOXML from TestRunner or API extensions |
| Keep **XLKit** as re-exports + thin `*+API` convenience | Grow a second parallel generation path beside `XLSXEngine` |
| Keep demo salts / passwords in **TestRunner** only | Add demo-only constants to public `CoreUtils` |

See ARCHITECTURE.md §2.2 for the full “where to put a change” table.

---

## 3. Excel / XLSX correctness

| Rule | Detail |
|------|--------|
| **`.xlsx` only** | Generate Office Open XML Excel workbooks. Do not claim `.numbers` or other formats. |
| **Column order** | Emit cells sorted by **numeric** column index (A…Z, AA…). Never rely on lexicographic string sort of addresses. |
| **Sheet visibility** | Emit `state` only for `.hidden` / `.veryHidden`. Emit `activeTab` when the first visible sheet is not index 0. |
| **Sheet protection** | Emit `<sheetProtection>` only when `Sheet.protection != nil`, and **after** `</sheetData>`. |
| **Password hashes** | Use `CoreUtils.excelLegacySheetPasswordHash` / `excelModernSheetPasswordHash` / `configureSheetPassword`. Do **not** use the incorrect OOXML-documented legacy formula. |
| **XML escaping** | All user strings in XML go through `CoreUtils.escapeXML`. |
| **Validate demos** | TestRunner generators that write `.xlsx` should validate with CoreXLSX where applicable. |

---

## 4. Image embedding (critical)

| Rule | Detail |
|------|--------|
| **ImageSizingUtils only** | All display size, Excel col/row, and EMU math goes through **ImageSizingUtils**. No magic numbers for sizing. |
| **Aspect ratio** | Preserve aspect ratio with zero intentional distortion. Do not stretch to fill cells. |
| **Formulas** | Column `pixels/8.0`, row `pixels/1.33`, EMU `pixels*9525`. |
| **Drawing offsets** | Prefer `rowOff = 0` / cell-aligned placement; do not “fix” mis-sizing with arbitrary offsets. |
| **Formats** | Support GIF / PNG / JPEG (JPG). Do not reintroduce BMP/TIFF without an explicit project decision. |
| **Workbook registration** | Auto-sized embed APIs must register images with **both** sheet and workbook when counting/saving depends on workbook images. |

---

## 5. Concurrency & platforms

| Rule | Detail |
|------|--------|
| **Not Sendable** | Do not mark `Workbook` / `Sheet` as Sendable. |
| **MainActor I/O** | Keep save, XLSX generation, SecurityManager, and image embed entry points on `@MainActor` unless a maintainer redesigns concurrency. |
| **Strict concurrency** | Library must build under `SWIFT_STRICT_CONCURRENCY=complete` (CI job). Prefer fixing isolation over `@unchecked Sendable`. |
| **iOS APIs** | Never use `FileManager.default.homeDirectoryForCurrentUser` without `#if os(macOS)`. Prefer `temporaryDirectory` / documents / caches. |
| **Supported platforms** | macOS 12+, iOS 15+. Do not add Linux/Windows/tvOS/watchOS targets without an explicit project decision. |

---

## 6. Tests & CLI

| Rule | Detail |
|------|--------|
| **Swift Testing** | Unit tests use **Swift Testing** (`@Suite`, `@Test`, `#expect` / `#require`). Do not reintroduce XCTest for new unit tests. |
| **Helpers** | Use **`XLKitTestSupport`** for dates, temp workbooks, and border helpers — do not subclass XCTestCase. |
| **Behaviour needs tests** | Public API and XLSX emission changes need unit tests (and CLI/CoreXLSX checks when demos are affected). |
| **Update counts** | Keep **80** (or the new total) aligned in `Tests/README.md`, Manual 10, AGENT.MD, `.cursorrules`, and this file when adding/removing tests. |
| **TestRunner ≠ unit tests** | CLI demos are not a substitute for `swift test`. |
| **Demo password** | Comprehensive demo password **1234** and salts stay in `ComprehensiveDemoProtection.swift` (TestRunner). |

---

## 7. Documentation & changelog

| Rule | Detail |
|------|--------|
| **AGENT ↔ .cursorrules** | When you update one, update the other. |
| **Feature docs** | User-visible behaviour → Manual (+ Test-Workflows / TestRunner README as needed). Structure → ARCHITECTURE.md. Hard constraints → this file. |
| **CHANGELOG** | Keep [Keep a Changelog](https://keepachangelog.com/) style; document user-visible and contributor-visible changes under Unreleased / version headings. |
| **Manual API chapter** | Public API additions/changes must update **Documentation/Manual/12-Complete-API-Reference.md** (and the relevant feature chapter). |

---

## 8. Safety & scope

| Rule | Detail |
|------|--------|
| **No unsafe / dynamic execution** | No unsafe pointers, `eval`-style reflection, or shelling out for generation. |
| **No exploit / malware work** | Do not write exploits, exploit PoCs, or attack tooling. |
| **Local files only** | Library file I/O stays local; no network fetch of workbooks or remote code execution in library targets. |
| **No secrets in git** | Do not commit `.env`, tokens, or private customer paths. |
| **Security features stay on** | Do not disable rate limiting / security logging / quarantine hooks without an explicit maintainer decision. Checksum storage may stay off by default. |
| **Dependencies** | Prefer pinned, reputable SPM deps in `Package.resolved`. Do not add heavy or unrelated packages casually. |
| **Destructive git only if asked** | No force-push to main, hard reset, or hook-skipping unless the user explicitly requests it. |
| **Commit only when asked** | Agents create commits only when the user explicitly requests a commit. |

---

## 9. Signs (learned constraints)

Append new signs when a failure repeats or a design decision must not drift. Keep each entry short.

### Template

```markdown
### Sign: short-title
- **Trigger:** When …
- **Instruction:** Always / Never …
- **Reason:** …
- **Provenance:** YYYY-MM-DD — brief note (PR / incident / design lock)
```

### Active signs

### Sign: image-sizing-utils-only
- **Trigger:** Changing how embedded images size cells or drawing XML.
- **Instruction:** Always route through `ImageSizingUtils`; never hardcode pixel→Excel or EMU conversions.
- **Reason:** Aspect-ratio regressions are highly visible; formulas are empirically locked to Excel behaviour.
- **Provenance:** 2025-07 — perfect aspect-ratio preservation design lock.

### Sign: numeric-column-sort
- **Trigger:** Emitting `<sheetData>` cells or validating column order.
- **Instruction:** Sort by numeric column index so AA follows Z; never lexicographic address sort.
- **Reason:** Excel repairs or rejects mis-ordered columns beyond Z.
- **Provenance:** 2025-10 — column ordering fix.

### Sign: legacy-password-not-ooxml-docs
- **Trigger:** Implementing or documenting worksheet protection password hashes.
- **Instruction:** Use `CoreUtils` Excel-compatible legacy + SHA-512 helpers; do not trust the OOXML-documented legacy formula.
- **Reason:** Incorrect hashes fail to unlock in Excel / LibreOffice.
- **Provenance:** 2026-06 — sheet password helpers (1.1.6).

### Sign: demo-salts-stay-in-testrunner
- **Trigger:** Need reproducible `Comprehensive-Demo.xlsx` modern hashes.
- **Instruction:** Keep salts/password in `ComprehensiveDemoProtection.swift`; never promote them to public CoreUtils.
- **Reason:** Demo fixtures must not pollute the library API.
- **Provenance:** 2026-06 — comprehensive demo protection design.

### Sign: workbook-sheet-not-sendable
- **Trigger:** Urge to “fix” concurrency by marking Workbook/Sheet Sendable or spawning Tasks over them.
- **Instruction:** Keep mutable model non-Sendable; keep save/generation on `@MainActor`.
- **Reason:** Shared mutable dictionaries; Sendable would be a lie.
- **Provenance:** Standing Swift 6 architecture decision.

### Sign: unit-tests-are-swift-testing
- **Trigger:** Adding or rewriting unit tests.
- **Instruction:** Use Swift Testing (`@Suite` / `@Test` / `#expect`); use `XLKitTestSupport` helpers.
- **Reason:** Full suite migrated; hybrid XCTest would confuse CI and docs.
- **Provenance:** 2026-07 — XCTest → Swift Testing migration.

---

## 10. Quick checklist before merge

- [ ] Change sits in the correct module (ARCHITECTURE §2.2 / Guardrails §2).
- [ ] Public behaviour has Swift Testing coverage; TestRunner/CoreXLSX updated if demos change.
- [ ] Image sizing uses ImageSizingUtils; column order is numeric.
- [ ] Sheet protection / visibility XML rules respected; password helpers are CoreUtils.
- [ ] Concurrency: no false Sendable; strict-concurrency CI still green.
- [ ] Docs: AGENT.MD ↔ .cursorrules if agent briefing changed; Manual / ARCHITECTURE / GUARDRAILS / CHANGELOG as needed.
- [ ] No secrets; security hooks not silently disabled.

---

## 11. References

- **Internal:** [ARCHITECTURE.md](ARCHITECTURE.md), [AGENT.MD](AGENT.MD), [.cursorrules](.cursorrules), [Tests/README.md](Tests/README.md), [Documentation/Manual/README.md](Documentation/Manual/README.md), [SECURITY.md](SECURITY.md).
- **External:** [Swift API Design Guidelines](https://swift.org/documentation/api-design-guidelines/), [Swift Testing](https://developer.apple.com/xcode/swift-testing/), [Swift Concurrency](https://docs.swift.org/swift-book/documentation/the-swift-programming-language/concurrency/).
