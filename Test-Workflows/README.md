# Test-Workflows

Excel workbooks produced by **`XLKitTestRunner`** for manual inspection, demos, and CI. Every generator that writes an `.xlsx` file validates structure with **CoreXLSX** where applicable.

Run all commands from the **package root** (`/path/to/XLKit`).

## Quick start

```bash
swift run XLKitTestRunner help
```

## Commands and output files

| Command | Aliases | Output file | Location |
|---------|---------|-------------|----------|
| `no-embeds` | `no-images` | `Embed-Test.xlsx` | This folder |
| `embed` | `with-embeds`, `with-images` | `Embed-Test-Embed.xlsx` | This folder |
| `comprehensive` | `demo` | `Comprehensive-Demo.xlsx` | This folder |
| `number-formats` | `formats`, `number-format` | `Number-Format-Test.xlsx` | This folder |
| `ios-test` | `ios` | `iOS-Example.xlsx` | Repository root |
| `security-demo` | `security` | *(console only)* | — |
| `sheet-password` | `password-hash`, `sheet-password-hash` | *(console only)* | — |

`sheet-password` does not create an Excel file; it prints hash values for worksheet protection (see below).

### Example invocations

```bash
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner number-formats
swift run XLKitTestRunner ios-test
swift run XLKitTestRunner security-demo

# Developer helper: worksheet protection hashes (no .xlsx)
swift run XLKitTestRunner sheet-password GreatDay
swift run XLKitTestRunner sheet-password 1234 --demo-salts
```

Requires **Swift 6** and dependencies resolved (`swift build` or `swift test` once). Test data for embed/no-embeds lives under **`Test-Data/Embed-Test/`**.

---

## Generated workbooks

### `Embed-Test.xlsx` (`no-embeds`)

- Built from **`Test-Data/Embed-Test/Embed-Test.csv`**
- CSV import, header formatting, auto column widths
- No embedded images

### `Embed-Test-Embed.xlsx` (`embed`)

- Same CSV plus PNG images from **`Test-Data/Embed-Test/`**
- Aspect-ratio preservation and image column sizing

### `Comprehensive-Demo.xlsx` (`comprehensive`)

Broad API sample: **11 sheets**, CoreXLSX validation by sheet name.

| Sheet | Purpose |
|-------|---------|
| Launch (Hidden) | First sheet `.hidden` — exercises `activeTab` on open |
| API Demo | Cells, formats, font colours, image, formulas |
| Second Sheet | Extra sheet |
| Fluent Demo | Chained row APIs, merge |
| Column Order Test | Columns A–AD (30 columns) |
| Reference (Hidden) | `.hidden` |
| Internal (Very Hidden) | `.veryHidden` |
| Protected (Basic) | Default protection, no password |
| Protected (Password) | Legacy + SHA-512; demo password **`1234`** |
| Protected (Granular) | Permission flags, no password |
| Protected (Modern Hash) | SHA-512 only; demo password **`1234`** |

**Unprotect in Excel**

| Sheet | Password |
|-------|----------|
| Protected (Basic) | *(none — leave blank)* |
| Protected (Password) | **`1234`** |
| Protected (Granular) | *(none)* |
| Protected (Modern Hash) | **`1234`** |

Open after generating:

```bash
swift run XLKitTestRunner comprehensive
open Test-Workflows/Comprehensive-Demo.xlsx
```

### `Number-Format-Test.xlsx` (`number-formats`)

Currency, percentage, and custom number formats.

### `iOS-Example.xlsx` (`ios-test`)

Written to the **repository root** (not this folder) using iOS-safe paths.

---

## `sheet-password` (developer utility)

Prints legacy and modern **`SheetProtection`** fields for a plaintext password. Use this when implementing protected sheets — do not guess `saltValue` / `hashValue`.

```bash
swift run XLKitTestRunner sheet-password YourPassword
```

- **Legacy:** four-character hex for `SheetProtection.password` (Excel still prompts for the **plaintext**).
- **Modern:** `algorithmName`, `saltValue`, `hashValue`, `spinCount` (random salt each run unless you fix salt in code).
- Prints a Swift snippet and suggests `CoreUtils.configureSheetPassword`.

For the comprehensive demo password and fixed salts:

```bash
swift run XLKitTestRunner sheet-password 1234 --demo-salts
```

**In application code**, prefer:

```swift
var protection = SheetProtection()
try CoreUtils.configureSheetPassword(&protection, plaintext: "YourPassword")
sheet.protection = protection
```

---

## Validation

Generators call **CoreXLSX** to confirm workbooks parse correctly (worksheets, shared strings, styles). Failures surface as errors in the CLI output.

---

## More documentation

- **`Sources/XLKitTestRunner/README.md`** — Full CLI reference, adding new test types, security notes  
- **`Documentation/Manual/10-Testing-Test-Runner-CI-and-Code-Style.md`** — Unit tests, CI workflows, code style  
- **`Tests/README.md`** — `XLKitTests` layout (80 unit tests)
