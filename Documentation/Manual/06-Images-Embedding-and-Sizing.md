# Chapter 06 — Images: Detection, Embedding, and Sizing

Navigation: [← Chapter 05](05-CSV-and-TSV.md) · [Manual index](README.md) · [Chapter 07 →](07-Formatting-Numbers-Alignment-and-Borders.md)

XLKit embeds **GIF**, **PNG**, and **JPEG** bytes into `.xlsx` media and drawing parts. The APIs you use most are on **`Sheet`**: **`embedImageAutoSized`** (recommended) and **`embedImage`**, plus file-based **`embedImage(from:)`**. Lower-level helpers live in **`ImageUtils`** and **`ImageSizingUtils`**.

Image APIs are **`@MainActor`** and use **`async throws`** (or `throws` for `ImageUtils` helpers that are synchronous). Call them from your main actor or an isolated async context.

---

## What “perfect sizing” means

- **Aspect ratio** is preserved: the bitmap is not stretched to arbitrary cell shapes.
- **`embedImageAutoSized`** computes a **display size** inside `maxCellWidth` × `maxCellHeight` (after an optional **scale**), then sets **column width** and **row height** using Excel-style conversions (`pixels / 8.0` for column width, `pixels / 1.33` for row height — see **`ImageSizingUtils`**).
- The XLSX engine writes **drawing** XML using **EMUs** (1 px ≈ 9525 EMUs). You normally do not call EMU helpers yourself unless you build custom layouts.

## Comprehensive aspect ratio support

**Comprehensive Aspect Ratio Support:** XLKit has been extensively tested with all professional video and cinema aspect ratios including:

- 16:9 (HD/4K video)
- 1:1 (Square format)
- 9:16 (Vertical video)
- 21:9 (Ultra-wide)
- 3:4 (Portrait)
- 2.39:1 (Cinemascope/Anamorphic)
- 1.85:1 (Academy ratio)
- 4:3 (Classic TV/monitor)
- 18:9 (Modern mobile)
- 1.19:1 (HD Standard)
- 1.5:1 (SD Academy)
- 1.48:1 (SD Academy Alt)
- 1.25:1 (SD Standard)
- 1.9:1 (IMAX Digital)
- 1.32:1 (DCI Standard)
- 2.37:1 (5K Cinema Scope)
- 1.37:1 (IMAX Film 15/70mm)

All aspect ratios are preserved with pixel-perfect accuracy using empirically derived Excel formulas, ensuring professional-quality exports for any video or cinema format. The API does not special-case ratio names—any image dimensions are scaled with a single **uniform** scale so **width:height** stays constant.

---

## Supported formats

| Format | Detection | Notes |
|--------|------------|--------|
| PNG | Magic bytes | Full support |
| JPEG | Magic bytes | `.jpeg` / `.jpg` |
| GIF | Magic bytes | Embedded frame; animated GIF is stored as GIF |

If detection fails, **`embedImageAutoSized`** returns **`false`** (or you get no image from **`ImageUtils`**).

---

## Recommended: `embedImageAutoSized`

Registers the image on the **sheet** and the **workbook**, sizes the row/column for the anchor cell, and returns **`true`** on success.

```swift
import XLKit

@MainActor
func addPhoto(data: Data, sheet: Sheet, workbook: Workbook) async throws {
    let ok = try await sheet.embedImageAutoSized(
        data,
        at: "B2",
        of: workbook,
        format: nil,              // nil = detect from bytes
        maxCellWidth: 600,
        maxCellHeight: 400,
        scale: 0.5               // 50% of max bounds after aspect fit
    )
    guard ok else {
        // Unsupported bytes or could not read dimensions
        return
    }
}
```

**Parameters (defaults):**

- **`maxCellWidth`** / **`maxCellHeight`**: upper bound in **points/pixels** for the fitted image (default **600 × 400**).
- **`scale`**: applied inside **`ImageSizingUtils.calculateDisplaySize`** (default **0.5**). Use **1.0** for the largest fit within the max box.

---

## Convenience: `embedImage` (scale wrapper)

**`embedImage`** multiplies **`maxWidth`** and **`maxHeight`** by **`scale`**, then forwards to **`embedImageAutoSized`**. Use it when you think in terms of “global scale” rather than editing max dimensions.

```swift
try await sheet.embedImage(
    imageData,
    at: "C3",
    of: workbook,
    scale: 0.7,
    maxWidth: 600,
    maxHeight: 400
)
```

---

## From disk: `embedImage(from:)` URL or path

These use **`ImageUtils.createExcelImage(from:displaySize:)`**, then **`sheet.addImage`**. They do **not** call **`workbook.addImage`** — the XLSX engine still collects images from sheets when building media, so exports work; if you maintain workbook-level image lists yourself, prefer **`embedImageAutoSized`** for a single code path.

```swift
let url = URL(fileURLWithPath: "/path/to/frame.png")
try await sheet.embedImage(from: url, at: "D1", displaySize: nil)

// String path overload
try await sheet.embedImage(from: "/path/to/frame.png", at: "E1")
```

---

## Manual `ExcelImage` + `addImage`

Use when you build **`Data`** and sizes yourself (e.g. tests or custom pipelines). You must place the image on the **sheet** and register it on the **workbook** if you rely on **`workbook.getImages()`**.

```swift
guard let format = ImageUtils.detectImageFormat(from: data),
      let size = ImageUtils.getImageSize(from: data, format: format) else { return }

let display = ImageSizingUtils.calculateDisplaySize(
    originalSize: size,
    maxWidth: 400,
    maxHeight: 300,
    scale: 1.0
)

let excelImage = ExcelImage(
    id: "image_\(UUID().uuidString)",
    data: data,
    format: format,
    originalSize: size,
    displaySize: display
)
sheet.addImage(excelImage, at: "F1")
workbook.addImage(excelImage)
```

**`ImageUtils.createExcelImage`** applies size limits and security checks; prefer it when creating **`ExcelImage`** from raw bytes in production code.

```swift
let image = try ImageUtils.createExcelImage(from: data, format: .png, displaySize: nil)
```

---

## Inspection helpers

```swift
if sheet.hasImage(at: "B2"), let img = sheet.getImage(at: "B2") {
    _ = img.originalSize
    _ = img.displaySize
}
sheet.removeImage(at: "B2")
```

---

## `ImageSizingUtils` (tooling and tests)

Use for custom layouts or unit tests — same helpers the engine uses for column/row sizing and EMUs:

```swift
let w = ImageSizingUtils.excelColumnWidth(forPixelWidth: 320)
let h = ImageSizingUtils.excelRowHeight(forPixelHeight: 240)
let emus = ImageSizingUtils.pixelsToEMUs(100)
```

See **[Chapter 12 — API reference](12-Complete-API-Reference.md)** for every **`ImageUtils`** / **`ImageSizingUtils`** symbol.
