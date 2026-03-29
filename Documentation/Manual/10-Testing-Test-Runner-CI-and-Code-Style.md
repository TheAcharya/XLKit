# Chapter 10 — Testing, XLKitTestRunner, CI, and Code Style

Navigation: [← Chapter 09](09-Errors-CoreUtils-and-iOS.md) · [Manual index](README.md) · [Chapter 11 →](11-Examples-Cookbook.md)

---

## Unit tests (`XLKitTests`)

The package test target **`XLKitTests`** exercises public APIs (workbook/sheet, CSV, images, formatting, merges, borders, file save, column ordering, etc.). Tests are **`@MainActor`** and inherit shared helpers from **`XLKitTestBase`**.

### Run everything locally

From the package root:

```bash
swift test
```

Run one test type (XCTest filter):

```bash
swift test --filter XLKitTests.CSVTests
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

    func testAsyncSave() async throws {
        try await withSavedTempWorkbookAsync(prefix: "my-async") { workbook, url in
            XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
        }
    }
}
```

### Deterministic dates

Use **`XLKitTestBase.makeUTCDate`** (or **`fixedTestDate`** / **`epochDate`**) instead of **`Date()`** when assertions depend on serialised values.

### Border helpers

**`XLKitTestBase`** exposes **`makeThinBorderFormat()`**, **`makeMediumRedBorderFormat()`**, **`makeThickBlueBorderFormat()`** for consistent border tests.

---

## `XLKitTestRunner` (CLI)

The executable **`XLKitTestRunner`** builds sample workbooks and (where noted) validates them with **CoreXLSX**. It is for **demos and CI smoke tests**, not a replacement for **`swift test`**.

### Commands

```bash
cd /path/to/XLKit   # package root

swift run XLKitTestRunner help

swift run XLKitTestRunner no-embeds      # CSV → xlsx, no images
swift run XLKitTestRunner embed          # CSV + embedded images
swift run XLKitTestRunner comprehensive  # broad API demo
swift run XLKitTestRunner security-demo  # path restriction demo
swift run XLKitTestRunner ios-test       # iOS-oriented sample output
swift run XLKitTestRunner number-formats # number format showcase
```

Aliases such as **`no-images`**, **`with-embeds`**, **`demo`**, **`formats`** are accepted — see **`Sources/XLKitTestRunner/main.swift`** for the full switch.

### Outputs

Generated files usually land under **`Test-Workflows/`** (or the project root for some scenarios). See **`Sources/XLKitTestRunner/README.md`** for paths per command.

### Add a new scenario

1. Copy **`Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift`** to a new file.
2. Implement your generator using **`XLKit`** APIs.
3. Register the subcommand in **`main.swift`** and extend **`printHelp()`**.

---

## CI (GitHub Actions)

Workflows live in **`.github/workflows/`**:

| Workflow | Role |
|----------|------|
| **`build.yml`** | macOS build + unit tests + **`swift run XLKitTestRunner`** smoke (`help`, `embed`); duplicate job with Swift tools 6.0; **iOS** simulator build + tests |
| **`codeql.yml`** | CodeQL security analysis |
| **`cli-*.yml`** | Focused CLI runs (e.g. no-embeds, embed, generic, iOS, number-formats) |

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

- **`Tests/README.md`** — Test layout and conventions  
- **`Sources/XLKitTestRunner/README.md`** — CLI details  
- **[Chapter 12 — API reference](12-Complete-API-Reference.md)** — Public surface under test  
