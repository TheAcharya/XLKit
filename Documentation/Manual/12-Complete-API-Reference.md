# Chapter 12 — Complete API Reference

Navigation: [← Chapter 11](11-Examples-Cookbook.md) · [Manual index](README.md)

This chapter lists the full public API of XLKit. All types and members are available when you `import XLKit`. For conceptual guides, see the earlier chapters.

### Workbook API

| Method / Property | Description |
|------------------|-------------|
| `init()` | Create an empty workbook. |
| `addSheet(name: String) -> Sheet` | Add a new sheet; returns the created sheet. |
| `getSheets() -> [Sheet]` | Return all sheets. |
| `getSheet(name: String) -> Sheet?` | Return the sheet with the given name. |
| `removeSheet(name: String)` | Remove the sheet with the given name. |
| `addImage(_ image: ExcelImage)` | Register an image with the workbook (required for XLSX export). |
| `getImages() -> [ExcelImage]` | Return all workbook-level images. |
| `getImage(withId id: String) -> ExcelImage?` | Return the image with the given id. |
| `removeImage(withId id: String)` | Remove the image with the given id. |
| `getImages(withFormat format: ImageFormat) -> [ExcelImage]` | Return images filtered by format. |
| `clearImages()` | Remove all workbook-level images. |
| `imageCount: Int` | Number of images in the workbook. |
| `init(fromCSV csvData: String, sheetName:hasHeader:)` | Convenience initializer: create a workbook from CSV (comma-separated). |
| `init(fromTSV tsvData: String, sheetName:hasHeader:)` | Convenience initializer: create a workbook from TSV. |
| `save(to url: URL) throws` | Save the workbook to a file (synchronous; `@MainActor`). |
| `save(to url: URL) async throws` | Save the workbook to a file (asynchronous; `@MainActor`). |
| `importCSV(_:into:hasHeader:)` | Import CSV into an existing sheet (comma-separated). |
| `importTSV(_:into:hasHeader:)` | Import TSV into an existing sheet. |
| `exportSheetToCSV(_ sheet:) -> String` | Export a sheet to CSV (comma-separated). |
| `exportSheetToTSV(_ sheet:) -> String` | Export a sheet to TSV. |

### Sheet API

**Cell and range operations**

| Method / Property | Description |
|------------------|-------------|
| `setCell(_ coordinate: String, value: CellValue) -> Self` | Set cell by value. |
| `setCell(_ coordinate: String, string: String, format: CellFormat? = nil) -> Self` | Set cell with string (and optional format). |
| `setCell(_ coordinate: String, number: Double, format: CellFormat? = nil) -> Self` | Set cell with number. |
| `setCell(_ coordinate: String, integer: Int, format: CellFormat? = nil) -> Self` | Set cell with integer. |
| `setCell(_ coordinate: String, boolean: Bool, format: CellFormat? = nil) -> Self` | Set cell with boolean. |
| `setCell(_ coordinate: String, date: Date, format: CellFormat? = nil) -> Self` | Set cell with date. |
| `setCell(_ coordinate: String, formula: String, format: CellFormat? = nil) -> Self` | Set cell with formula. |
| `setCell(_ coordinate: String, cell: Cell) -> Self` | Set cell with a `Cell` (value + format). |
| `setCell(row: Int, column: Int, value: CellValue) -> Self` | Set cell by row/column. |
| `setCell(row: Int, column: Int, cell: Cell) -> Self` | Set cell by row/column with `Cell`. |
| `getCell(_ coordinate: String) -> CellValue?` | Get cell value. |
| `getCellWithFormat(_ coordinate: String) -> Cell?` | Get cell value and format. |
| `getCellFormat(_ coordinate: String) -> CellFormat?` | Get cell format only. |
| `setRange(_ range: String, cell: Cell) -> Self` | Set a range with a `Cell`. |
| `setRange(_ range: String, string: String, format: CellFormat? = nil) -> Self` | Set range with string. |
| `setRange(_ range: String, number: Double, format: CellFormat? = nil) -> Self` | Set range with number. |
| `setRange(_ range: String, integer: Int, format: CellFormat? = nil) -> Self` | Set range with integer. |
| `setRange(_ range: String, boolean: Bool, format: CellFormat? = nil) -> Self` | Set range with boolean. |
| `setRange(_ range: String, date: Date, format: CellFormat? = nil) -> Self` | Set range with date. |
| `setRange(_ range: String, formula: String, format: CellFormat? = nil) -> Self` | Set range with formula. |
| `mergeCells(_ range: String) -> Self` | Merge the given range. |
| `getUsedCells() -> [String]` | Return sorted list of coordinates that have values. |
| `getMergedRanges() -> [CellRange]` | Return all merged ranges. |
| `setRow(_ row: Int, values: [CellValue]) -> Self` | Set a row with an array of values. |
| `setRow(_ row: Int, strings: [String]) -> Self` | Set a row with strings. |
| `setRow(_ row: Int, numbers: [Double]) -> Self` | Set a row with numbers. |
| `setRow(_ row: Int, integers: [Int]) -> Self` | Set a row with integers. |
| `setColumn(_ column: Int, values: [CellValue]) -> Self` | Set a column with values. |
| `setColumn(_ column: Int, strings: [String]) -> Self` | Set a column with strings. |
| `setColumn(_ column: Int, numbers: [Double]) -> Self` | Set a column with numbers. |
| `setColumn(_ column: Int, integers: [Int]) -> Self` | Set a column with integers. |

**Column and row sizing**

| Method / Property | Description |
|------------------|-------------|
| `setColumnWidth(_ column: Int, width: Double) -> Self` | Set column width (Excel units). |
| `getColumnWidth(_ column: Int) -> Double?` | Get column width. |
| `getColumnWidths() -> [Int: Double]` | Get all column widths. |
| `setRowHeight(_ row: Int, height: Double) -> Self` | Set row height. |
| `getRowHeight(_ row: Int) -> Double?` | Get row height. |
| `getRowHeights() -> [Int: Double]` | Get all row heights. |
| `autoSizeColumn(_ column: Int, forImageAt coordinate: String) -> Self` | Set column width to fit the image at the coordinate. |

**Image operations**

| Method / Property | Description |
|------------------|-------------|
| `addImage(_ image: ExcelImage, at coordinate: String) -> Self` | Place an image at a cell. |
| `addImage(_ image: ExcelImage, at row: Int, column: Int) -> Self` | Place an image by row/column. |
| `addImage(_ data: Data, at coordinate: String, format: ImageFormat, displaySize: CGSize? = nil) throws -> Bool` | Create and add image from `Data` (Sheet extension from XLKitImages). |
| `addImage(from url: URL, at coordinate: String, displaySize: CGSize? = nil) throws -> Bool` | Create and add image from file URL (Sheet extension). |
| `getImages() -> [String: ExcelImage]` | Return all images in the sheet (coordinate → image). |
| `getImage(at coordinate: String) -> ExcelImage?` | Get image at a coordinate. |
| `getImage(at row: Int, column: Int) -> ExcelImage?` | Get image by row/column. |
| `removeImage(at coordinate: String) -> Self` | Remove image at coordinate. |
| `removeImage(at row: Int, column: Int) -> Self` | Remove image by row/column. |
| `hasImage(at coordinate: String) -> Bool` | Whether the cell has an image. |
| `hasImage(at row: Int, column: Int) -> Bool` | Whether the cell has an image (by row/column). |
| `embedImageAutoSized(_ data:at:of:format:maxCellWidth:maxCellHeight:scale:) async throws -> Bool` | Embed image with automatic sizing and aspect ratio preservation; registers with workbook. |
| `embedImage(_ data:at:of:scale:maxWidth:maxHeight:) async throws -> Bool` | Embed image with scaling parameters. |
| `embedImage(from url: URL, at coordinate: String, displaySize: CGSize? = nil) async throws -> Bool` | Embed image from file URL. |
| `embedImage(from path: String, at coordinate: String, displaySize: CGSize? = nil) async throws -> Bool` | Embed image from file path. |
| `exportToCSV() -> String` | Export sheet to CSV (comma-separated). |
| `exportToTSV() -> String` | Export sheet to TSV. |

**Other**

| Method / Property | Description |
|------------------|-------------|
| `state: SheetState` | Sheet visibility in the tab bar (`.visible`, `.hidden`, `.veryHidden`). Default: `.visible`. |
| `protection: SheetProtection?` | Per-sheet protection settings; `nil` (default) means unprotected. |
| `clear()` | Clear all cells, formats, images, merged ranges, and dimensions (does not reset `state` or `protection`). |
| `allCells: [String: CellValue]` | All cells with values. |
| `allFormattedCells: [String: Cell]` | All cells with value and format. |
| `isEmpty: Bool` | Whether the sheet has any content. |
| `cellCount: Int` | Number of cells with values. |
| `imageCount: Int` | Number of images in the sheet. |
| `name: String` | Sheet name. |
| `id: Int` | Sheet id. |
| `cells`, `mergedRanges`, `columnWidths`, `rowHeights`, `images`, `cellFormats` | Internal storage (public for access). |

### SheetState (enum)

| Case | Description |
|------|-------------|
| `.visible` | Sheet shown in tab bar (default); no `state` attribute in workbook XML. |
| `.hidden` | Hidden from tab bar; user can unhide in Excel. Emits `state="hidden"`. |
| `.veryHidden` | Hidden and not unhideable from Excel UI. Emits `state="veryHidden"`. |

### SheetProtection (struct)

Maps 1:1 to the `<sheetProtection>` XLSX element. Boolean flags use **inverted lock semantics** — `true` means the action is locked when the sheet is protected.

| Property | Description |
|----------|-------------|
| `sheet: Bool?` | Whether protection is enforced. Default: `true`. |
| `password: String?` | Legacy 16-bit password hash. |
| `algorithmName: String?` | Modern hash algorithm (e.g. `"SHA-512"`). |
| `hashValue: String?` | Modern password hash. |
| `saltValue: String?` | Salt for modern hash. |
| `spinCount: Int?` | Iteration count for modern hash. |
| `objects: Bool?` | Lock drawing objects. XLSX default: `false`. |
| `scenarios: Bool?` | Lock scenario edits. XLSX default: `false`. |
| `formatCells: Bool?` | Lock cell formatting. XLSX default: `true`. |
| `formatColumns: Bool?` | Lock column formatting. XLSX default: `true`. |
| `formatRows: Bool?` | Lock row formatting. XLSX default: `true`. |
| `insertColumns: Bool?` | Lock column insertion. XLSX default: `true`. |
| `insertRows: Bool?` | Lock row insertion. XLSX default: `true`. |
| `insertHyperlinks: Bool?` | Lock hyperlink insertion. XLSX default: `true`. |
| `deleteColumns: Bool?` | Lock column deletion. XLSX default: `true`. |
| `deleteRows: Bool?` | Lock row deletion. XLSX default: `true`. |
| `selectLockedCells: Bool?` | Block selection of locked cells. XLSX default: `false`. |
| `selectUnlockedCells: Bool?` | Block selection of unlocked cells. XLSX default: `false`. |
| `sort: Bool?` | Lock sort operations. XLSX default: `true`. |
| `autoFilter: Bool?` | Lock AutoFilter. XLSX default: `true`. |
| `pivotTables: Bool?` | Lock PivotTable operations. XLSX default: `true`. |
| `init()` | Create with defaults (`sheet: true`, all permission flags `nil`). |

### Cell & coordinate types

**CellCoordinate**

| Member | Description |
|--------|-------------|
| `init(row: Int, column: Int)` | Create from 1-based row and column. |
| `init?(excelAddress: String)` | Parse Excel address (e.g. `"A1"`, `"AA10"`); case-insensitive. |
| `row: Int` | 1-based row. |
| `column: Int` | 1-based column (1 = A). |
| `excelAddress: String` | Excel-style address (e.g. `"A1"`). |

**CellRange**

| Member | Description |
|--------|-------------|
| `init(start: CellCoordinate, end: CellCoordinate)` | Create from start and end coordinates. |
| `init?(excelRange: String)` | Parse range (e.g. `"A1:B5"`). |
| `start: CellCoordinate`, `end: CellCoordinate` | Bounds. |
| `coordinates: [CellCoordinate]` | All coordinates in the range. |
| `excelRange: String` | Excel-style range (e.g. `"A1:B5"`). |

**CellValue** (enum)

Cases: `.string(String)`, `.number(Double)`, `.integer(Int)`, `.boolean(Bool)`, `.date(Date)`, `.formula(String)`, `.empty`.  
Properties: `stringValue: String`, `type: String`.

**Cell** (struct)

| Member | Description |
|--------|-------------|
| `init(_ value: CellValue, format: CellFormat? = nil)` | Create from value and optional format. |
| `value: CellValue`, `format: CellFormat?` | Stored value and format. |
| `static func string(_ value: String, format: CellFormat? = nil) -> Cell` | Factory for string cell. |
| `static func number(_ value: Double, format: CellFormat? = nil) -> Cell` | Factory for number cell. |
| `static func integer(_ value: Int, format: CellFormat? = nil) -> Cell` | Factory for integer cell. |
| `static func boolean(_ value: Bool, format: CellFormat? = nil) -> Cell` | Factory for boolean cell. |
| `static func date(_ value: Date, format: CellFormat? = nil) -> Cell` | Factory for date cell. |
| `static func formula(_ value: String, format: CellFormat? = nil) -> Cell` | Factory for formula cell. |

**CellFormat** (struct)

Properties: `fontName`, `fontSize`, `fontWeight`, `fontStyle`, `textDecoration`, `fontColor`, `backgroundColor`, `horizontalAlignment`, `verticalAlignment`, `textWrapping`, `textRotation`, `numberFormat`, `customNumberFormat`, `borderTop`, `borderBottom`, `borderLeft`, `borderRight`, `borderColor` (all optional where applicable).

Static factories: `text(fontName:fontSize:fontWeight:fontStyle:color:)`, `header(fontSize:backgroundColor:)`, `currency(format:color:)`, `percentage(format:)`, `date(format:)`, `bordered(style:color:)`.

**ExcelImage** (struct)

`id: String`, `data: Data`, `format: ImageFormat`, `originalSize: CGSize`, `displaySize: CGSize?`.  
`init(id:data:format:originalSize:displaySize:)`.

### Enums

- **SheetState**: `.visible`, `.hidden`, `.veryHidden`
- **FontWeight**: `.normal`, `.bold`
- **FontStyle**: `.normal`, `.italic`
- **TextDecoration**: `.none`, `.underline`, `.strikethrough`, `.underlineStrikethrough`
- **HorizontalAlignment**: `.left`, `.center`, `.right`, `.justify`, `.distributed`
- **VerticalAlignment**: `.top`, `.center`, `.bottom`, `.justify`, `.distributed`
- **BorderStyle**: `.none`, `.thin`, `.medium`, `.thick`, `.double`, `.hair`, `.dotted`, `.dashed`, `.dashDot`, `.dashDotDot`, `.slantDashDot`
- **NumberFormat**: `.general`, `.number`, `.numberWithDecimals`, `.percentage`, `.percentageWithDecimals`, `.currency`, `.currencyWithDecimals`, `.date`, `.dateTime`, `.time`, `.text`, `.custom`; `static func custom(_ format: String) -> NumberFormat`
- **ImageFormat**: `.gif`, `.png`, `.jpeg`, `.jpg`; properties: `mimeType`, `excelContentType`

### CoreUtils

Static methods and constants (XLKitCore):

| Member | Description |
|--------|-------------|
| `columnLetter(from column: Int) -> String` | Convert column number to letter (1 → `"A"`). |
| `columnNumber(from letter: String) -> Int` | Convert letter to number (`"A"` → 1). |
| `formatDate(_ date: Date) -> String` | Format date for Excel (yyyy-mm-dd). |
| `dateFromExcelNumber(_ number: Double) -> Date` | Convert Excel date number to `Date`. |
| `excelNumberFromDate(_ date: Date) -> Double` | Convert `Date` to Excel date number. |
| `escapeXML(_ string: String) -> String` | Escape XML special characters. |
| `xmlHeader() -> String` | Standard XML declaration string. |
| `validateFilePath(_ path: String) throws` | Validate path for security (`@MainActor`). |
| `validateFileSize(_ data: Data, maxSize: Int) throws` | Ensure data size ≤ max. |
| `validateImageFileSize(_ data: Data) throws` | Validate against image size limit. |
| `validateExcelFileSize(_ data: Data) throws` | Validate against Excel file size limit. |
| `generateChecksum(_ data: Data) -> String` | SHA-256 hex string. |
| `generateFileChecksum(_ fileURL: URL) throws -> String` | SHA-256 of file contents. |
| `safeFileURL(for filename: String) -> URL` | Platform-safe URL (e.g. documents dir on iOS). |
| `excelLegacySheetPasswordHash(for password: String) -> String` | Legacy 16-bit hex for `SheetProtection.password`. |
| `excelModernSheetPasswordHash(for:spinCount:salt:) throws -> ExcelModernSheetPasswordHash` | SHA-512 salt/hash/spinCount for Excel 2013+. |
| `configureSheetPassword(_:plaintext:legacy:modern:spinCount:salt:) throws` | Fill `SheetProtection` from plaintext password. |
| `excelModernSheetPasswordDefaultSpinCount` | Default `100_000` spin count. |
| `maxImageFileSize: Int` | 50 MB. |
| `maxExcelFileSize: Int` | 100 MB. |
| `allowedDirectories: [URL]` | Allowed base directories for path validation. |

Demo password and salts for **`Comprehensive-Demo.xlsx`** live in **`Sources/XLKitTestRunner/ComprehensiveDemoProtection.swift`** (TestRunner only, not exported by the `XLKit` library). See [Chapter 10](10-Testing-Test-Runner-CI-and-Code-Style.md) and **`Test-Workflows/README.md`**.

### SecurityManager API

Static configuration and methods (`@MainActor`):

| Member | Description |
|--------|-------------|
| `enableChecksumStorage: Bool` | When true, store/verify file checksums (default: false). |
| `enableFilePathRestrictions: Bool` | When true, restrict saves to allowed directories (default: false). |
| `checkRateLimit() throws` | Enforce rate limit (e.g. 100 ops/min); throws if exceeded. |
| `logSecurityOperation(_ operation: String, details: [String: Any])` | Log a security-relevant event. |
| `quarantineSuspiciousFile(_ fileURL: URL, reason: String) throws` | Move file to quarantine and throw. |
| `shouldQuarantineFile(_ data: Data, format: ImageFormat) -> Bool` | Whether data/file should be quarantined. |
| `storeFileChecksum(_ checksum: String, for fileURL: URL)` | Store checksum (if checksum storage enabled). |
| `verifyFileChecksum(_ fileURL: URL) -> Bool` | Verify file against stored checksum. |

### ImageUtils & ImageSizingUtils

**ImageUtils** (static; `@MainActor` where used):

| Member | Description |
|--------|-------------|
| `detectImageFormat(from data: Data) -> ImageFormat?` | Detect format from magic bytes (GIF, PNG, JPEG). |
| `getImageSize(from data: Data, format: ImageFormat) -> CGSize?` | Read dimensions from image data. |
| `createExcelImage(from data: Data, format: ImageFormat, displaySize: CGSize? = nil) throws -> ExcelImage?` | Create `ExcelImage` with validation and quarantine check. |
| `createExcelImage(from url: URL, displaySize: CGSize? = nil) throws -> ExcelImage?` | Load from URL and create `ExcelImage`. |

**ImageSizingUtils** (static; used internally and for custom sizing):

| Member | Description |
|--------|-------------|
| `calculateDisplaySize(originalSize:maxWidth:maxHeight:minWidth:minHeight:scale:) -> CGSize` | Scale size within bounds preserving aspect ratio. |
| `excelColumnWidth(forPixelWidth pixelWidth: CGFloat) -> CGFloat` | Pixels to Excel column width (pixels / 8.0). |
| `excelRowHeight(forPixelHeight pixelHeight: CGFloat) -> CGFloat` | Pixels to Excel row height (pixels / 1.33). |
| `pixelsToEMUs(_ pixels: CGFloat) -> Int64` | Convert pixels to EMUs (Excel drawing units). |
| `emusToPixels(_ emus: Int64) -> CGFloat` | Convert EMUs to pixels. |
| `calculateDrawingDimensions(imageSize:maxWidth:maxHeight:) -> (width: Int64, height: Int64)` | Drawing size in EMUs. |
| `idealCellSizeForImage(imageWidth:imageHeight:) -> (colWidth: CGFloat, rowHeight: CGFloat)` | Excel column width and row height to fit image. |
| `cellPixelSize(colWidth:rowHeight:) -> (width: CGFloat, height: CGFloat)` | Pixel size from Excel column/row dimensions. |
| `imageOffsetsInCell(imageWidth:imageHeight:cellWidth:cellHeight:) -> (xEMU: Int64, yEMU: Int64)` | Offsets to center image in cell (EMUs). |
| `drawingEMUs(imageWidth:imageHeight:) -> (cx: Int64, cy: Int64)` | Drawing dimensions in EMUs. |

### CSVUtils

Static methods; typically used via `Workbook`/`Sheet` instance methods. Only CSV (comma) and TSV (tab) are supported. Parsing and generation are powered by the [swift-textfile](https://github.com/orchetect/swift-textfile) library for spec-compliant handling of quoted fields, escaped quotes, and edge cases (e.g. RFC 4180 for CSV).

| Member | Description |
|--------|-------------|
| `exportToCSV(sheet: Sheet) -> String` | Export sheet to CSV (comma-separated; spec-compliant via TextFile). |
| `exportToTSV(sheet: Sheet) -> String` | Export sheet to TSV (tab-separated; spec-compliant via TextFile). |
| `importFromCSV(sheet: Sheet, csvData: String, hasHeader: Bool)` | Import CSV into sheet (comma-separated). |
| `importFromTSV(sheet: Sheet, tsvData: String, hasHeader: Bool)` | Import TSV into sheet. |
| `createWorkbookFromCSV(csvData: String, sheetName: String, hasHeader: Bool) -> Workbook` | New workbook from CSV (comma-separated). |
| `createWorkbookFromTSV(tsvData: String, sheetName: String, hasHeader: Bool) -> Workbook` | New workbook from TSV. |

### XLKitError

Cases: `invalidCoordinate(String)`, `invalidRange(String)`, `fileWriteError(String)`, `zipCreationError(String)`, `xmlGenerationError(String)`, `securityError(String)`, `rateLimitExceeded(String)`, `fileSizeLimitExceeded(String)`, `suspiciousFileDetected(String)`.  
Conforms to `Error` and `LocalizedError`; `errorDescription: String?`.

### XLSXEngine

| Member | Description |
|--------|-------------|
| `generateXLSX(workbook: Workbook, to url: URL) throws` | Generate and write the .xlsx file (`@MainActor`). |
| `formatToKey(_ format: CellFormat) -> String` | Serialize format to a key (for internal/style mapping). |

Typically you call `workbook.save(to: url)` rather than `XLSXEngine.generateXLSX` directly.


### ImageSizingUtils (additional members)

| Method | Description |
|--------|-------------|
| `calculatePositioningOffsets(imageSize:cellWidth:cellHeight:) -> (x: Int64, y: Int64)` | EMU offsets for centering (advanced layout). |
| `calculateRowOffset(imageHeight:) -> Int64` | Fixed row offset constant used in drawing calculations. |

### SecurityManager.SecurityLogEntry

| Property | Description |
|----------|-------------|
| `timestamp: Date` | Log time. |
| `operation: String` | Operation name. |
| `details: [String: Any]` | Structured metadata. |
| `userAgent: String` | Client identifier string. |

### XLKit namespace

| Declaration | Description |
|-------------|-------------|
| `public struct XLKit { }` | Empty namespace; re-exports attach APIs to `Workbook` and `Sheet`. |

