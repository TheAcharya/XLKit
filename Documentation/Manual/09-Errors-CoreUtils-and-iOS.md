# Chapter 09 — Errors, CoreUtils, and iOS

Navigation: [← Chapter 08](08-Security-and-Validation.md) · [Manual index](README.md) · [Chapter 10 →](10-Testing-Test-Runner-CI-and-Code-Style.md)

---

## `XLKitError`

All failure cases used by XLKit conform to **`XLKitError`**. Prefer pattern matching on the enum so you preserve associated messages.

### Saving a workbook

```swift
import XLKit

@MainActor
func exportWorkbook(_ workbook: Workbook, to url: URL) async {
    do {
        try await workbook.save(to: url)
    } catch let error as XLKitError {
        switch error {
        case .securityError(let message):
            print("Path not allowed: \(message)")
        case .rateLimitExceeded(let message):
            print("Slow down: \(message)")
        case .fileWriteError(let message), .zipCreationError(let message), .xmlGenerationError(let message):
            print("Export failed: \(message)")
        default:
            print("Other XLKit error: \(error.localizedDescription)")
        }
    } catch {
        print("System error: \(error)")
    }
}
```

### Images and security

Image loading can throw **`fileSizeLimitExceeded`**, **`suspiciousFileDetected`**, or **`securityError`** (paths) depending on **`SecurityManager`** / **`CoreUtils`** settings:

```swift
do {
    if let img = try ImageUtils.createExcelImage(from: data, format: .png) {
        _ = img.id // place with `sheet.addImage(_:at:)` / `workbook.addImage` as needed
    }
} catch let error as XLKitError {
    if case .suspiciousFileDetected(let msg) = error {
        print(msg)
    }
} catch {
    print(error)
}
```

### Localized descriptions

**`XLKitError`** conforms to **`LocalizedError`**; you can show **`error.localizedDescription`** in UI.

Full case list: **[Chapter 12 — API reference](12-Complete-API-Reference.md)** (`XLKitError`).

---

## `CoreUtils` usage

### Column letters and Excel addresses

```swift
let letters = CoreUtils.columnLetter(from: 28)   // "AB"
let n = CoreUtils.columnNumber(from: "AB")     // 28
let coord = CellCoordinate(row: 5, column: 3) // C5
print(coord.excelAddress)
```

### Dates and Excel serial numbers

Cells store dates as **numbers** in the worksheet XML; XLKit converts using an Excel 1900-era epoch:

```swift
let someDate = Date() // or your fixed test date from Calendar
let excelSerial = CoreUtils.excelNumberFromDate(someDate)
let roundTrip = CoreUtils.dateFromExcelNumber(excelSerial)
let label = CoreUtils.formatDate(roundTrip) // yyyy-mm-dd style string for stringValue paths
```

### XML escaping

If you generate XML beside XLKit, reuse:

```swift
let safe = CoreUtils.escapeXML("A & B < C")
```

### Checksums (files)

```swift
if let sha = try? CoreUtils.generateFileChecksum(fileURL) {
    print(sha)
}
let dataChecksum = CoreUtils.generateChecksum(Data("hello".utf8))
```

### Size limits

```swift
do {
    try CoreUtils.validateImageFileSize(imageData)
    try CoreUtils.validateExcelFileSize(xlsxData)
} catch let error as XLKitError {
    if case .fileSizeLimitExceeded(let msg) = error { print(msg) }
}
```

### Worksheet protection passwords

Do not invent `saltValue` or `hashValue` strings. Use the public helpers on `CoreUtils` (see [Chapter 12 — CoreUtils](12-Complete-API-Reference.md)):

```swift
var protection = SheetProtection()
try CoreUtils.configureSheetPassword(&protection, plaintext: "mySecret")
sheet.protection = protection
```

- **Legacy only (16-bit hex):** `protection.password = CoreUtils.excelLegacySheetPasswordHash(for: "mySecret")` — Excel still prompts for the **plaintext** when unprotecting.
- **Modern (SHA-512, Excel 2013+):** `try CoreUtils.excelModernSheetPasswordHash(for: "mySecret")` returns `algorithmName`, `saltValue`, `hashValue`, and `spinCount` (default spin count `100_000`). Pass an optional `salt: Data` when you need a stable hash across runs.
- **CLI helper (no `.xlsx` written):** `swift run XLKitTestRunner sheet-password mySecret` — see [Chapter 10 — `sheet-password`](10-Testing-Test-Runner-CI-and-Code-Style.md).

Demo password **`1234`** and salts for **`Comprehensive-Demo.xlsx`** are defined in **`ComprehensiveDemoProtection.swift`** inside **XLKitTestRunner** only (not part of the library product).

### File paths (optional restrictions)

When **`SecurityManager.enableFilePathRestrictions`** is **`true`**, **`CoreUtils.validateFilePath`** runs on save paths. On **iOS**, validation is relaxed when restrictions are on (sandbox enforces access); on **macOS**, paths must sit under allowed roots — see [Chapter 08](08-Security-and-Validation.md).

```swift
do {
    try CoreUtils.validateFilePath(url.path)
} catch let error as XLKitError {
    if case .securityError(let msg) = error { print(msg) }
}
```

---

## iOS: saving where the sandbox allows

Use a URL under the app’s **Documents**, **Caches**, or **temporary** directory — not arbitrary POSIX paths.

### Recommended helper

**`CoreUtils.safeFileURL(for:)`** returns a file URL under **Documents** on iOS (and a sensible location on macOS for the helper’s contract — see source if you rely on macOS behaviour).

```swift
import XLKit

@MainActor
func saveForIOS(_ workbook: Workbook) async throws {
    let url = CoreUtils.safeFileURL(for: "report.xlsx")
    try await workbook.save(to: url)
}
```

### Explicit Documents directory

```swift
if let documents = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first {
    let url = documents.appendingPathComponent("report.xlsx")
    try await workbook.save(to: url)
}
```

### Info.plist (Files app / sharing)

To expose exports in **Files** or **document browser** workflows, enable the usual keys (exact keys vary by app requirements):

```xml
<key>UIFileSharingEnabled</key>
<true/>
<key>LSSupportsOpeningDocumentsInPlace</key>
<true/>
```

### Main actor

**`Workbook.save`**, **`SecurityManager`**, and image helpers are **`@MainActor`**. Declare your export UI or service as **`@MainActor`**, or hop to the main actor before calling XLKit:

```swift
await MainActor.run {
    try? await workbook.save(to: url)
}
```

---

## Quick map

| Need | API |
|------|-----|
| User-visible error text | `XLKitError.localizedDescription` |
| Column index ↔ letters | `CoreUtils.columnLetter`, `columnNumber` |
| Date ↔ Excel serial | `excelNumberFromDate`, `dateFromExcelNumber` |
| Safe output URL on iOS | `CoreUtils.safeFileURL(for:)` |
| Sheet protection from plaintext | `configureSheetPassword`, `excelLegacySheetPasswordHash`, `excelModernSheetPasswordHash` |
| SHA-256 | `generateChecksum`, `generateFileChecksum` |
