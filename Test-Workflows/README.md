# Test-Workflows

Excel workbooks produced by **`XLKitTestRunner`** for manual inspection, demos, and CI. Generators that write an `.xlsx` validate structure with **CoreXLSX** where applicable.

Run commands from the **package root** (`/path/to/XLKit`).

## Quick start

```bash
swift run XLKitTestRunner help
swift run XLKitTestRunner comprehensive
open Test-Workflows/Comprehensive-Demo.xlsx
```

Requires **Swift 6** and resolved dependencies (`swift build` or `swift test` once).

---

## Commands and output files

| Command | Aliases | Output | Location |
|---------|---------|--------|----------|
| `no-embeds` | `no-images` | `Embed-Test.xlsx` | This folder |
| `embed` | `with-embeds`, `with-images` | `Embed-Test-Embed.xlsx` | This folder |
| `comprehensive` | `demo` | `Comprehensive-Demo.xlsx` | This folder |
| `number-formats` | `formats`, `number-format` | `Number-Format-Test.xlsx` | This folder |
| `ios-test` | `ios` | `iOS-Example.xlsx` | Repository root |
| `security-demo` | `security` | *(console only)* | — |
| `sheet-password` | `password-hash`, `sheet-password-hash` | *(console only)* | — |

```bash
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner number-formats
swift run XLKitTestRunner ios-test
swift run XLKitTestRunner security-demo

swift run XLKitTestRunner sheet-password GreatDay
swift run XLKitTestRunner sheet-password 1234 --demo-salts
```

Embed/no-embeds input data: **`Test-Data/Embed-Test/`** (`Embed-Test.csv` and PNG files).

---

## Generated workbooks

### `Embed-Test.xlsx` (`no-embeds`)

- CSV import from **`Test-Data/Embed-Test/Embed-Test.csv`**
- Header formatting and auto column widths
- No embedded images

### `Embed-Test-Embed.xlsx` (`embed`)

- Same CSV plus PNG images from **`Test-Data/Embed-Test/`**
- Aspect-ratio preservation and image column sizing

### `Comprehensive-Demo.xlsx` (`comprehensive`)

11-sheet API sample; CLI validates every sheet by name after save.

| Sheet | Purpose |
|-------|---------|
| Launch (Hidden) | First sheet `.hidden` — `activeTab` opens **API Demo** |
| API Demo | Cells, formats, font colours, embedded image, formulas |
| Second Sheet | Additional sheet |
| Fluent Demo | Chained row APIs and merge |
| Column Order Test | Columns A–AD (30 columns, beyond Z) |
| Reference (Hidden) | `.hidden` |
| Internal (Very Hidden) | `.veryHidden` |
| Protected (Basic) | `SheetProtection()` with no password |
| Protected (Password) | Legacy hash + SHA-512 (password **`1234`**) |
| Protected (Granular) | Permission flags only, no password |
| Protected (Modern Hash) | SHA-512 only (password **`1234`**, different salt than Password sheet) |

**Unprotect in Excel**

| Sheet | Password | Notes |
|-------|----------|--------|
| Protected (Basic) | *(blank)* | Sheet is protected; unprotect without a password |
| Protected (Password) | **`1234`** | Legacy `CC3D` + modern hash via `configureSheetPassword` |
| Protected (Granular) | *(blank)* | After unprotect: format cells and insert rows allowed; sort stays locked |
| Protected (Modern Hash) | **`1234`** | Modern hash only (no legacy `password` attribute) |

Demo password and salts are in **`Sources/XLKitTestRunner/ComprehensiveDemoProtection.swift`** (TestRunner only, not XLKitCore). Regenerate after changing that file or protection logic:

```bash
swift run XLKitTestRunner comprehensive
```

### `Number-Format-Test.xlsx` (`number-formats`)

Currency, percentage, and custom number formats.

### `iOS-Example.xlsx` (`ios-test`)

Written to the **repository root** (not this folder) using iOS-safe paths.

### `security-demo` (`security`)

No file output. Demonstrates path validation, security logging, and rate limiting in the console.

---

## `sheet-password` (developer utility)

Prints legacy and modern **`SheetProtection`** fields for a plaintext password. Does not write an `.xlsx`.

```bash
swift run XLKitTestRunner sheet-password YourPassword
```

| Output | Meaning |
|--------|---------|
| Legacy hex | `SheetProtection.password` (Excel still asks for the **plaintext** when unprotecting) |
| Modern fields | `algorithmName`, `saltValue`, `hashValue`, `spinCount` |
| Swift snippet | Suggests `CoreUtils.configureSheetPassword` |

Each run uses a **random** modern salt unless you pass `salt:` in code. For **`Comprehensive-Demo.xlsx`** with password **`1234`**:

```bash
swift run XLKitTestRunner sheet-password 1234 --demo-salts
```

**Application code** (preferred over copying hash strings):

```swift
var protection = SheetProtection()
try CoreUtils.configureSheetPassword(&protection, plaintext: "YourPassword")
sheet.protection = protection
```

See **`Documentation/Manual/09-Errors-CoreUtils-and-iOS.md`** (worksheet protection passwords).

---

## CI

GitHub Actions run selected CLI commands and upload artifacts:

| Workflow | Commands | Artifact |
|----------|----------|----------|
| `cli-no-embeds.yml` | `no-embeds` | `Embed-Test.xlsx` |
| `cli-embed.yml` | `embed` | `Embed-Test-Embed.xlsx` |
| `cli-generic.yml` | `comprehensive`, `security-demo` | `Comprehensive-Demo.xlsx` |
| `cli-numbers.yml` | `number-formats` | `Number-Format-Test.xlsx` |
| `cli-ios.yml` | `ios-test` | `iOS-Example.xlsx` (root) |

`build.yml` also smoke-runs `help` and `embed`, and includes a **macOS (strict concurrency)** build/test job (`SWIFT_STRICT_CONCURRENCY=complete`). Details: **`Documentation/Manual/10-Testing-Test-Runner-CI-and-Code-Style.md`**.

---

## Validation

Generators call **CoreXLSX** to confirm workbooks parse (worksheets, shared strings, styles). Failures appear as errors in CLI output and fail CI.

---

## More documentation

- **`Sources/XLKitTestRunner/README.md`** — Full CLI reference and adding new commands  
- **`Documentation/Manual/10-Testing-Test-Runner-CI-and-Code-Style.md`** — Unit tests, CI, code style  
- **`Documentation/Manual/03-Core-Model-Workbook-Sheet-and-Cells.md`** — `SheetState`, `SheetProtection`  
- **`Tests/README.md`** — `XLKitTests` (80 unit tests)
