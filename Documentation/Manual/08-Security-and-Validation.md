# Chapter 08 — Security and Validation

Navigation: [← Chapter 07](07-Formatting-Numbers-Alignment-and-Borders.md) · [Manual index](README.md) · [Chapter 09 →](09-Errors-CoreUtils-and-iOS.md)

XLKit integrates **`SecurityManager`** (rate limiting, logging, optional checksum sidecars, optional path rules) and **`CoreUtils`** validation (image size, paths when restrictions are on). Most hooks run automatically when you save workbooks or process images; you only need the APIs below when you change defaults or build custom tooling.

`SecurityManager` is **`@MainActor`** — call it from the main actor (same as `Workbook.save`).

---

## Automatic behaviour (no code required)

| Trigger | What happens |
|--------|----------------|
| **`workbook.save(to:)`** | `SecurityManager.checkRateLimit()` (max **100** save operations per **60** seconds), `CoreUtils.validateFilePath` if path restrictions are enabled, structured `logSecurityOperation` for start/complete, optional checksum file next to the `.xlsx` if checksum storage is on |
| **`ImageUtils.createExcelImage`** | Image size cap, `shouldQuarantineFile` check, security logging |

You do not call `checkRateLimit()` yourself for normal saves; **`XLSXEngine`** does that internally.

---

## Configuration

### Optional checksum files

After each successful save, a SHA-256 hash can be written beside the file as `yourfile.xlsx.checksum` when enabled:

```swift
import XLKit

SecurityManager.enableChecksumStorage = true

let workbook = Workbook()
let sheet = workbook.addSheet(name: "Data")
sheet.setCell("A1", value: .string("Hello"))

let url = FileManager.default.temporaryDirectory.appendingPathComponent("report.xlsx")
try workbook.save(to: url)

// If enableChecksumStorage is true, report.xlsx.checksum may exist; verify later:
if SecurityManager.enableChecksumStorage {
    let ok = SecurityManager.verifyFileChecksum(url)
    print("Checksum matches: \(ok)")
}
```

With **`enableChecksumStorage == false`** (default), `verifyFileChecksum` returns **`true`** without reading a sidecar.

### Optional file path restrictions

When **`enableFilePathRestrictions`** is **`true`**, **`CoreUtils.validateFilePath`** enforces allowed base directories (temporary + home on macOS; iOS uses sandbox-friendly rules). Default is **`false`** (save anywhere the process can write).

```swift
SecurityManager.enableFilePathRestrictions = true

let url = FileManager.default.temporaryDirectory.appendingPathComponent("safe.xlsx")
let workbook = Workbook()
workbook.addSheet(name: "S").setCell("A1", value: .string("x"))
try workbook.save(to: url) // OK under tmp

// Saving outside allowed roots throws XLKitError.securityError on macOS when restrictions are on.
```

---

## Logging and custom events

Use **`logSecurityOperation(_:details:)`** for audit-style events in your own app (console + append to `TemporaryDirectory/xlkit_security_logs/security.log`):

```swift
SecurityManager.logSecurityOperation("export_started", details: [
    "user_id": "anonymous",
    "sheet_count": workbook.getSheets().count,
])
```

---

## Image quarantine (advanced)

**`ImageUtils.createExcelImage`** already calls **`shouldQuarantineFile`** and throws **`XLKitError.suspiciousFileDetected`** when data matches suspicious UTF-8 substrings or exceeds per-format size caps. You can mirror the check before loading untrusted data:

```swift
let data = // ...
if let format = ImageUtils.detectImageFormat(from: data),
   SecurityManager.shouldQuarantineFile(data, format: format) {
    // Reject or handle without embedding
}
```

**`quarantineSuspiciousFile(_:reason:)`** moves a file into `TemporaryDirectory/xlkit_quarantine/` and throws **`suspiciousFileDetected`** — use for file-URL workflows, not raw `Data`.

---

## Rate limiting

**`checkRateLimit()`** is invoked automatically inside **`workbook.save`**. Call it yourself only for **other** heavy file work you want to subject to the same limit (100 operations per rolling 60 seconds), so you do not double-count saves:

```swift
// Example: your own export loop that does not use workbook.save every iteration
for item in items {
    try SecurityManager.checkRateLimit()
    try writeAuxiliaryFile(for: item)
}
```

If the limit is exceeded, **`XLKitError.rateLimitExceeded`** is thrown.

---

## Integration summary

| Component | Role |
|-----------|------|
| **XLSXEngine** | Rate limit, path validation, logging, checksum store after ZIP write |
| **ImageUtils** | Quarantine heuristic, size validation, logging |
| **CoreUtils** | `validateFilePath`, `validateImageFileSize`, checksum helpers |

---

## Full API

See **[Chapter 12 — Complete API reference](12-Complete-API-Reference.md)** (`SecurityManager`, `CoreUtils`).
