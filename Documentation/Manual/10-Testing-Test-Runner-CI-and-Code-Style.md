# Chapter 10 — Testing, XLKitTestRunner, CI, and Code Style

Navigation: [← Chapter 09](09-Errors-CoreUtils-and-iOS.md) · [Manual index](README.md) · [Chapter 11 →](11-Examples-Cookbook.md)

---

## Unit tests (`XLKitTests`)

The package test target **`XLKitTests`** exercises public APIs (workbook/sheet, CSV, images, formatting, merges, borders, sheet state, sheet protection, file save, column ordering, etc.). Tests are **`@MainActor`** and inherit shared helpers from **`XLKitTestBase`**.

### Run everything locally

From the package root:

```bash
swift test
```

Run one test class (XCTest filter):

```bash
swift test --filter XLKitTests.SheetProtectionTests
```

### Pattern: temporary workbook on disk

Use **`withSavedTempWorkbookSync`** or **`withSavedTempWorkbookAsync`** when the test must read the generated **`.xlsx`** (or assert file existence). The helper saves first, runs your closure, then deletes the temp file.

```swift
import XCTest
import XLKit

@MainActor
final class MyFeatureTests: XLKitTestBase {
    func testSaveProducesFile() throws {
        try withSavedTempWorkbookSync(prefix: "my-case") { workbook, url in
            XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
            XCTAssertEqual(workbook.getSheets().count, 1)
        }
    }
}
```

### Deterministic dates

Use **`XLKitTestBase.makeUTCDate`** (or **`fixedTestDate`** / **`epochDate`**) instead of **`Date()`** when assertions depend on serialised values.

### Sheet state and protection tests

| File | Tests | Coverage |
|------|-------|----------|
| **`SheetStateTests.swift`** | 7 | `.visible` / `.hidden` / `.veryHidden`, workbook `state` and `activeTab` XML, save round-trip |
| **`SheetProtectionTests.swift`** | 14 | Legacy/modern password hashes, `configureSheetPassword`, `<sheetProtection>` XML, save round-trip |

The suite currently has **80 tests** across **15** focused files (see **`Tests/README.md`**).

---

## `XLKitTestRunner` (CLI)

The executable **`XLKitTestRunner`** builds sample workbooks and (where noted) validates them with **CoreXLSX**. Use it for **demos, manual QA, and CI smoke tests** — not as a substitute for **`swift test`**.

### Run from the package root

```bash
cd /path/to/XLKit

swift run XLKitTestRunner help
```

### Generate Excel files

| Command | Aliases | Output |
|---------|---------|--------|
| `no-embeds` | `no-images` | `Test-Workflows/Embed-Test.xlsx` |
| `embed` | `with-embeds`, `with-images` | `Test-Workflows/Embed-Test-Embed.xlsx` |
| `comprehensive` | `demo` | `Test-Workflows/Comprehensive-Demo.xlsx` |
| `number-formats` | `formats` | `Test-Workflows/Number-Format-Test.xlsx` |
| `ios-test` | `ios` | `iOS-Example.xlsx` (repo root) |
| `security-demo` | `security` | Console only (path restrictions) |

```bash
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner number-formats
swift run XLKitTestRunner ios-test
swift run XLKitTestRunner security-demo
```

**`comprehensive`** produces an 11-sheet workbook (formatting, images, CSV patterns, column ordering beyond Z, sheet visibility, sheet protection). Password-protected demo sheets use **`1234`** — see **`Test-Workflows/README.md`** for the sheet list and unprotect passwords.

Embed/no-embeds tests read CSV and images from **`Test-Data/Embed-Test/`**.

### Worksheet protection helper (`sheet-password`)

Developer utility: prints legacy hash and SHA-512 fields for a plaintext password (no `.xlsx` written).

```bash
swift run XLKitTestRunner sheet-password GreatDay
swift run XLKitTestRunner sheet-password 1234 --demo-salts
```

| Flag / input | Behaviour |
|--------------|-----------|
| `<password>` argument | Required plaintext (e.g. `GreatDay`) |
| `--demo-salts` | When password is **`1234`**, also prints salts from `ComprehensiveDemoProtection.swift` (same as **`Comprehensive-Demo.xlsx`**) |

Modern `saltValue` / `hashValue` are **random each run** unless you supply a fixed `salt` in code. In apps, use:

```swift
var protection = SheetProtection()
try CoreUtils.configureSheetPassword(&protection, plaintext: "GreatDay")
sheet.protection = protection
```

See [Chapter 03 — Sheet visibility and protection](03-Core-Model-Workbook-Sheet-and-Cells.md).

### Output locations

| Path | Contents |
|------|----------|
| **`Test-Workflows/`** | Most generated `.xlsx` files — see **`Test-Workflows/README.md`** |
| **Repository root** | `iOS-Example.xlsx` from `ios-test` |

### Add a new CLI scenario

1. Copy **`Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift`** to a new generator file.
2. Implement using **`XLKit`** APIs; validate with **CoreXLSX** when writing `.xlsx`.
3. Register the subcommand in **`Sources/XLKitTestRunner/main.swift`** and extend **`printHelp()`**.
4. Document output path in **`Test-Workflows/README.md`** and **`Sources/XLKitTestRunner/README.md`**.

### Source layout

| File | Role |
|------|------|
| `main.swift` | CLI switch and help |
| `ExcelGenerators.swift` | CSV, comprehensive, security, iOS, number formats |
| `ImageEmbedGenerators.swift` | Image embed workflow |
| `SheetPasswordUtilities.swift` | `sheet-password` output |
| `ComprehensiveDemoProtection.swift` | Sheet protection constants for the `comprehensive` CLI generator |
| `Templates/TestGeneratorTemplate.swift` | Starter for new commands |

---

## CI (GitHub Actions)

Workflows live in **`.github/workflows/`**:

| Workflow | Role |
|----------|------|
| **`build.yml`** | macOS build + unit tests + **`swift run XLKitTestRunner`** smoke (`help`, `embed`); **iOS** simulator build + tests |
| **`codeql.yml`** | CodeQL security analysis |
| **`cli-generic.yml`** | `comprehensive`, `security-demo`; uploads **`Comprehensive-Demo.xlsx`** |
| **`cli-embed.yml`** / **`cli-no-embeds.yml`** | Image embed and CSV-only workflows |
| **`cli-ios.yml`** | iOS-oriented CLI run |

Pushes that touch only docs/assets may skip builds — see each workflow’s **`paths-ignore`**.

---

## Code style

Project formatting expectations are described in **`.cursorrules`** and enforced via **`swift-format`** using **`.swift-format`** (4 spaces, 120 columns, grouped imports, trailing commas, etc.).

Format sources from the repo root (with [swift-format](https://github.com/swiftlang/swift-format) installed):

```bash
swift-format format --configuration .swift-format -i .
```

Use the same configuration in Xcode or CI so diffs stay consistent.

---

## Where to read more

- **`Test-Workflows/README.md`** — Generated files, comprehensive demo sheets, passwords  
- **`Sources/XLKitTestRunner/README.md`** — Full CLI reference  
- **`Tests/README.md`** — Unit test layout  
- **[Chapter 12 — API reference](12-Complete-API-Reference.md)** — Public surface under test  
