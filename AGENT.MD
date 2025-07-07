# XLKit - AI Agent Development Guide

## Library Overview

XLKit is a modern Swift library for creating and manipulating Excel (.xlsx) files on macOS. The library is built with a modular architecture using Swift Package Manager (SPM) modules, targeting Swift 6.0 and macOS 14+.

### Core Architecture

The library is organized into the following SPM modules:

1. **XLKitCore** - Core types, data structures, and utilities
2. **XLKitFormatters** - CSV/TSV import/export functionality
3. **XLKitImages** - Image processing and embedding utilities
4. **XLKitXLSX** - XLSX file generation engine
5. **XLKit** - Main API that re-exports all submodules

### Module Dependencies

```
XLKit (main API)
├── XLKitCore (core types & utilities)
├── XLKitFormatters (CSV/TSV import/export)
├── XLKitImages (image processing & embedding)
└── XLKitXLSX (XLSX generation engine)

XLKitFormatters
└── XLKitCore

XLKitImages
└── XLKitCore

XLKitXLSX
├── XLKitCore
├── XLKitFormatters
└── XLKitImages
```

### Executable Target

```
XLKitTestRunner (executable)
└── XLKit
```

## XLKitTestRunner

### Overview

XLKitTestRunner is a modular command-line tool for generating Excel files for testing and demonstration purposes. It provides a structured way to create various types of Excel files to test XLKit functionality.

### Architecture

**File Structure**:
```
Sources/XLKitTestRunner/
├── main.swift                    # Entry point with CLI
├── ExcelGenerators.swift         # Excel generation functions
├── Templates/                    # Template files
│   └── TestGeneratorTemplate.swift
└── README.md                     # Documentation
```

**Key Components**:

1. **main.swift**: Command-line interface that routes to different test functions
2. **ExcelGenerators.swift**: Contains the actual Excel generation logic
3. **Templates/**: Reusable templates for creating new tests
4. **README.md**: Comprehensive documentation

### Usage

**Basic Commands**:
```bash
# Run default test (no-embeds)
swift run XLKitTestRunner

# Run specific test
swift run XLKitTestRunner no-embeds

# Show help
swift run XLKitTestRunner help
```

**Available Test Types**:
- `no-embeds` / `no-images`: Generate Excel from CSV without images
- `with-embeds` / `with-images`: Generate Excel with image embeds (future)
- `csv-import`: Test CSV import functionality (future)
- `formatting`: Test cell formatting features (future)
- `embed`: Generate Excel with embedded images from CSV data

### Development Workflow

**Adding New Tests**:

1. **Use Template**: Copy the template file
   ```bash
   cp Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift Sources/XLKitTestRunner/YourTestName.swift
   ```

2. **Modify Template**: Update function name, configuration, and logic

3. **Register Test**: Add to main.swift switch statement
   ```swift
   case "your-test-name":
       print("Executing: Your Test Description")
       yourTestName()
   ```

4. **Update Help**: Add to help function in main.swift

5. **Create Workflow**: Add GitHub Actions workflow if needed

**Naming Conventions**:
- **Function Names**: camelCase (e.g., `generateExcelWithImages()`)
- **Test Types**: kebab-case (e.g., `with-images`, `csv-import`)
- **File Names**: PascalCase (e.g., `ExcelGenerators.swift`)

### Output Management

**Generated Files**: All Excel files are saved to `Test-Workflows/` directory
- `Embed-Test.xlsx`: From no-embeds test
- `Embed-Test-Embed.xlsx`: From embed test (with images)
- `[Your-Test].xlsx`: From custom tests

**Error Handling**: All functions use consistent error handling with descriptive messages and appropriate exit codes.

### Integration

**GitHub Actions**: The test runner is integrated into CI/CD workflows
- `.github/workflows/create-excel-test-file-no-images.yml`: Runs `no-embeds` test
- Uses `xcbeautify` for clean output
- Generates artifacts for download

**Testing Strategy**: Provides a way to test XLKit functionality in real-world scenarios by generating actual Excel files.

## Detailed Module Breakdown

### XLKitCore Module

**File**: `Sources/XLKitCore/CoreTypes.swift`

**Purpose**: Contains all core data structures and types used throughout the library.

#### Key Types

**Workbook Class**
```swift
public final class Workbook {
    private var sheets: [Sheet] = []
    private let nextSheetId: Int
    private var images: [ExcelImage] = []
}
```
- Manages multiple sheets and workbook-level images
- Auto-increments sheet IDs
- Note: Not Sendable due to mutable state requirements

**Sheet Class**
```swift
public final class Sheet: Equatable {
    public let name: String
    public let id: Int
    private var cells: [String: CellValue] = [:]
    private var mergedRanges: [CellRange] = []
    private var columnWidths: [Int: Double] = [:]
    private var rowHeights: [Int: Double] = [:]
    private var images: [String: ExcelImage] = [:]
    private var cellFormats: [String: CellFormat] = [:]
}
```
- Represents a worksheet with cells, formatting, and images
- Supports cell operations, range operations, and image embedding
- Maintains separate storage for cell values and formats
- Note: Not Sendable due to mutable state requirements

**CellValue Enum**
```swift
public enum CellValue: Equatable {
    case string(String)
    case number(Double)
    case integer(Int)
    case boolean(Bool)
    case date(Date)
    case formula(String)
    case empty
}
```
- Type-safe representation of cell content
- Supports all Excel data types
- Includes computed properties for string representation and type identification

**CellCoordinate Struct**
```swift
public struct CellCoordinate: Hashable {
    public let row: Int
    public let column: Int
    
    public var excelAddress: String // e.g., "A1", "B2"
}
```
- Represents Excel-style coordinates (row, column)
- Converts between numeric and Excel address notation
- Supports parsing Excel addresses like "A1", "B2", "AA10"

**CellRange Struct**
```swift
public struct CellRange: Hashable {
    public let start: CellCoordinate
    public let end: CellCoordinate
    
    public var coordinates: [CellCoordinate] // All coordinates in range
    public var excelRange: String // e.g., "A1:B3"
}
```
- Represents ranges of cells
- Generates all coordinates within the range
- Supports Excel range notation like "A1:B3"

**Cell Struct**
```swift
public struct Cell {
    public let value: CellValue
    public let format: CellFormat?
    
    // Factory methods for each cell type
    public static func string(_ value: String, format: CellFormat? = nil) -> Cell
    public static func number(_ value: Double, format: CellFormat? = nil) -> Cell
    public static func integer(_ value: Int, format: CellFormat? = nil) -> Cell
    public static func boolean(_ value: Bool, format: CellFormat? = nil) -> Cell
    public static func date(_ value: Date, format: CellFormat? = nil) -> Cell
    public static func formula(_ value: String, format: CellFormat? = nil) -> Cell
}
```
- Combines cell value with optional formatting
- Provides factory methods for easy cell creation

**CellFormat Struct**
```swift
public struct CellFormat {
    public var fontName: String?
    public var fontSize: Double?
    public var fontWeight: FontWeight?
    public var fontStyle: FontStyle?
    public var fontColor: String?
    public var backgroundColor: String?
    public var horizontalAlignment: HorizontalAlignment?
    public var verticalAlignment: VerticalAlignment?
    public var textDecoration: TextDecoration?
    public var textWrapping: Bool?
    public var textRotation: Int?
    public var numberFormat: NumberFormat?
    public var customNumberFormat: String?
    public var borderTop: BorderStyle?
    public var borderBottom: BorderStyle?
    public var borderLeft: BorderStyle?
    public var borderRight: BorderStyle?
    public var borderColor: String?
    
    // Predefined format factories
    public static func header(fontSize: Double = 14.0, backgroundColor: String = "#E0E0E0") -> CellFormat
    public static func currency(format: NumberFormat = .currencyWithDecimals, color: String? = nil) -> CellFormat
    public static func percentage(format: NumberFormat = .percentageWithDecimals) -> CellFormat
    public static func date(format: NumberFormat = .date) -> CellFormat
    public static func bordered(style: BorderStyle = .thin, color: String? = nil) -> CellFormat
}
```
- Comprehensive cell formatting options
- Supports fonts, colors, alignment, borders, and number formats
- Provides factory methods for common formatting patterns

**CoreUtils Struct**
```swift
public struct CoreUtils {
    // Column conversion
    public static func columnLetter(from column: Int) -> String
    public static func columnNumber(from letter: String) -> Int
    
    // Date conversion
    public static func formatDate(_ date: Date) -> String
    public static func dateFromExcelNumber(_ number: Double) -> Date
    public static func excelNumberFromDate(_ date: Date) -> Double
    
    // XML utilities
    public static func escapeXML(_ string: String) -> String
    public static func xmlHeader() -> String
}
```
- Utility functions for common operations
- Handles Excel-specific conversions (column letters, dates)
- Provides XML escaping for XLSX generation

**ExcelImage Struct**
```swift
public struct ExcelImage {
    public let id: String
    public let data: Data
    public let format: ImageFormat
    public let originalSize: CGSize
    public let displaySize: CGSize?
}
```
- Represents an image to be embedded in Excel
- Stores image data, format, and size information
- Supports custom display sizes

**ImageFormat Enum**
```swift
public enum ImageFormat: String, CaseIterable {
    case gif = "gif"
    case png = "png"
    case jpeg = "jpeg"
    case jpg = "jpg"
    case bmp = "bmp"
    case tiff = "tiff"
    
    public var mimeType: String
    public var excelContentType: String
}
```
- Defines supported image formats
- Provides MIME types and Excel content types

**XLKitError Enum**
```swift
public enum XLKitError: Error, LocalizedError {
    case invalidCoordinate(String)
    case invalidRange(String)
    case fileWriteError(String)
    case zipCreationError(String)
    case xmlGenerationError(String)
}
```
- Comprehensive error types for all operations
- Provides localized error descriptions

### XLKitFormatters Module

**File**: `Sources/XLKitFormatters/CSVUtils.swift`

**Purpose**: Handles CSV and TSV import/export operations.

**CSVUtils Struct**
```swift
public struct CSVUtils {
    // Export functions
    public static func exportToCSV(sheet: Sheet, separator: String = ",") -> String
    public static func exportToTSV(sheet: Sheet) -> String
    
    // Import functions
    public static func importFromCSV(sheet: Sheet, csvData: String, separator: String = ",", hasHeader: Bool = false)
    public static func importFromTSV(sheet: Sheet, tsvData: String, hasHeader: Bool = false)
    
    // Workbook creation from CSV/TSV
    public static func createWorkbookFromCSV(csvData: String, sheetName: String = "Sheet1", separator: String = ",", hasHeader: Bool = false) -> Workbook
    public static func createWorkbookFromTSV(tsvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) -> Workbook
}
```

**Key Features**:
- Handles CSV/TSV parsing with proper quote escaping
- Supports header row handling (skips first row when hasHeader = true)
- Auto-detects data types (string, number, boolean, date)
- Handles special characters and newlines in CSV data
- Creates workbooks directly from CSV/TSV data

**Implementation Details**:
- Uses regex for date pattern matching
- Implements proper CSV parsing with quote handling
- Supports both comma and tab separators
- Maintains data integrity during import/export

### XLKitImages Module

**File**: `Sources/XLKitImages/ImageUtils.swift`

**Purpose**: Handles image processing, format detection, and size extraction.

**ImageUtils Struct**
```swift
public struct ImageUtils {
    // Format detection
    public static func detectImageFormat(from data: Data) -> ImageFormat?
    
    // Size detection
    public static func getImageSize(from data: Data, format: ImageFormat) -> CGSize?
    
    // ExcelImage creation
    public static func createExcelImage(from data: Data, format: ImageFormat, displaySize: CGSize? = nil) -> ExcelImage?
    public static func createExcelImage(from url: URL, displaySize: CGSize? = nil) throws -> ExcelImage?
}
```

**Supported Formats**:
- **GIF**: Detects signature `GIF`, extracts width/height from header
- **PNG**: Detects signature `89 50 4E 47`, extracts dimensions from IHDR chunk
- **JPEG**: Detects signature `FF D8`, parses SOF markers for dimensions
- **BMP**: Detects signature `42 4D`, extracts dimensions from header
- **TIFF**: Detects signature `49 49` or `4D 4D`, simplified size detection

**Implementation Details**:
- Binary signature detection for format identification
- Format-specific size extraction algorithms
- Error handling for corrupted or unsupported images
- Support for both Data and URL input sources

### XLKitXLSX Module

**File**: `Sources/XLKitXLSX/XLSXEngine.swift`

**Purpose**: Generates XLSX files using the OpenXML format.

**XLSXEngine Struct**
```swift
public struct XLSXEngine {
    // Async and sync save operations
    public static func generateXLSX(workbook: Workbook, to url: URL) async throws
    public static func generateXLSXSync(workbook: Workbook, to url: URL) throws
}
```

**XLSX Generation Process**:
1. **Content Types**: Generates `[Content_Types].xml` with file type mappings
2. **Workbook XML**: Creates `xl/workbook.xml` with sheet definitions
3. **Worksheets**: Generates individual `xl/worksheets/sheetN.xml` files
4. **Relationships**: Creates relationship files for proper file linking
5. **ZIP Archive**: Packages all XML files into a ZIP archive

**File Structure Generated**:
```
workbook.xlsx
├── [Content_Types].xml
├── _rels/
│   └── .rels
└── xl/
    ├── workbook.xml
    ├── _rels/
    │   └── workbook.xml.rels
    └── worksheets/
        ├── sheet1.xml
        └── _rels/
            └── sheet1.xml.rels
```

**Implementation Details**:
- Uses temporary directories for file generation
- Implements proper XML escaping
- Handles ZIP creation using system `zip` command
- Supports both synchronous and asynchronous operations
- Includes proper cleanup on errors

### XLKit Module (Main API)

**File**: `Sources/XLKit/XLKit.swift`

**Purpose**: Provides the main public API that re-exports all submodule functionality.

**XLKit Struct**
```swift
public struct XLKit {
    // Workbook creation
    public static func createWorkbook() -> Workbook
    
    // Save operations
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) async throws
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) throws
    
    // CSV/TSV convenience methods
    public static func createWorkbookFromCSV(csvData: String, sheetName: String = "Sheet1", separator: String = ",", hasHeader: Bool = false) -> Workbook
    public static func createWorkbookFromTSV(tsvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) -> Workbook
    public static func exportSheetToCSV(sheet: Sheet, separator: String = ",") -> String
    public static func exportSheetToTSV(sheet: Sheet) -> String
    public static func importCSVIntoSheet(sheet: Sheet, csvData: String, separator: String = ",", hasHeader: Bool = false)
    public static func importTSVIntoSheet(sheet: Sheet, tsvData: String, hasHeader: Bool = false)
}
```

**API Design Principles**:
- Static methods for easy access
- Fluent API design with method chaining
- Comprehensive error handling
- Type-safe operations
- Zero external dependencies

## Sheet API Reference

### Cell Operations

```swift
// Basic cell setting
sheet.setCell("A1", value: .string("Hello"))
sheet.setCell("B1", value: .number(42.5))
sheet.setCell("C1", value: .integer(100))
sheet.setCell("D1", value: .boolean(true))
sheet.setCell("E1", value: .date(Date()))
sheet.setCell("F1", value: .formula("=A1+B1"))

// Cell setting with formatting
sheet.setCell("A1", cell: Cell.string("Hello", format: CellFormat.header()))

// Cell setting by row/column
sheet.setCell(row: 1, column: 1, value: .string("A1"))

// Cell retrieval
let value = sheet.getCell("A1")
let cellWithFormat = sheet.getCellWithFormat("A1")
let format = sheet.getCellFormat("A1")
```

### Range Operations

```swift
// Set range with same value
sheet.setRange("A1:C3", value: .string("Range"))

// Set range with cell (includes formatting)
sheet.setRange("A1:C3", cell: Cell.string("Range", format: CellFormat.bordered()))

// Merge cells
sheet.mergeCells("A1:C1")

// Get merged ranges
let mergedRanges = sheet.getMergedRanges()
```

### Column and Row Operations

```swift
// Column width
sheet.setColumnWidth(1, width: 100.0)
let width = sheet.getColumnWidth(1)
let allWidths = sheet.getColumnWidths()

// Row height
sheet.setRowHeight(1, height: 50.0)
let height = sheet.getRowHeight(1)
let allHeights = sheet.getRowHeights()

// Auto-size column for image
sheet.autoSizeColumn(2, forImageAt: "B2")
```

### Image Operations

```swift
// Add image from data
sheet.addImage(imageData, at: "B2", format: .png)

// Add image from URL
try sheet.addImage(imageURL, at: "B2")

// Add ExcelImage object
sheet.addImage(excelImage, at: "B2")

// Get all images
let images = sheet.getImages()

// Embed image with automatic sizing (preserves aspect ratio)
try sheet.embedImageAutoSized(
    imageData: imageData,
    at: "B2",
    workbook: workbook
)

// Embed image with automatic sizing from URL
try sheet.embedImageAutoSized(
    imageURL: imageURL,
    at: "B2",
    workbook: workbook
)

// XLKit convenience method for auto-sized embedding
try XLKit.embedImageAutoSized(
    in: sheet,
    imageData: imageData,
    at: "B2",
    workbook: workbook
)
```

**Image Embedding Features**:
- **Aspect Ratio Preservation**: Images maintain their original proportions
- **Automatic Cell Sizing**: Row height and column width automatically adjust to match image dimensions
- **Precise Positioning**: Images are positioned exactly at cell boundaries with no offsets
- **Excel Compliance**: Uses precise Excel formulas for sizing calculations
- **Workbook Registration**: Images are registered with both sheet and workbook for proper tracking

### Utility Operations

```swift
// Get used cells
let usedCells = sheet.getUsedCells()

// Clear all data
sheet.clear()
```

## Workbook API Reference

### Sheet Management

```swift
// Add sheets
let sheet1 = workbook.addSheet(name: "Sheet1")
let sheet2 = workbook.addSheet(name: "Sheet2")

// Get sheets
let allSheets = workbook.getSheets()
let specificSheet = workbook.getSheet(name: "Sheet1")

// Remove sheets
workbook.removeSheet(name: "Sheet1")
```

### Image Management

```swift
// Add workbook-level images
workbook.addImage(excelImage)

// Get all images
let images = workbook.getImages()
```

## CSV/TSV API Reference

### Import Operations

```swift
// Create workbook from CSV
let workbook = XLKit.createWorkbookFromCSV(
    csvData: "Name,Age\nJohn,25\nJane,30",
    sheetName: "Data",
    separator: ",",
    hasHeader: true
)

// Import into existing sheet
XLKit.importCSVIntoSheet(
    sheet: sheet,
    csvData: "Name,Age\nJohn,25\nJane,30",
    separator: ",",
    hasHeader: true
)
```

### Export Operations

```swift
// Export sheet to CSV
let csv = XLKit.exportSheetToCSV(sheet: sheet, separator: ",")

// Export sheet to TSV
let tsv = XLKit.exportSheetToTSV(sheet: sheet)
```

## Error Handling

### XLKitError Types

```swift
// Coordinate errors
XLKitError.invalidCoordinate("Z999999") // Invalid Excel address

// Range errors
XLKitError.invalidRange("A1:Z999999") // Invalid range

// File operation errors
XLKitError.fileWriteError("Permission denied")
XLKitError.zipCreationError("ZIP creation failed")
XLKitError.xmlGenerationError("XML generation failed")
```

### Error Handling Patterns

```swift
do {
    try XLKit.saveWorkbook(workbook, to: url)
} catch XLKitError.invalidCoordinate(let coord) {
    print("Invalid coordinate: \(coord)")
} catch XLKitError.fileWriteError(let message) {
    print("File write error: \(message)")
} catch {
    print("Unexpected error: \(error)")
}
```

## Testing Strategy

### Test Coverage

The library includes comprehensive tests in `Tests/XLKitTests/XLKitTests.swift` covering:

1. **Core Functionality**
   - Workbook creation and management
   - Sheet operations
   - Cell value setting and retrieval
   - Coordinate and range operations

2. **CSV/TSV Operations**
   - Import with and without headers
   - Export with various separators
   - Special character handling
   - Data type detection

3. **Image Operations**
   - Format detection
   - Size extraction
   - ExcelImage creation
   - Image embedding in sheets

4. **XLSX Generation**
   - File creation and saving
   - XML structure validation
   - ZIP archive creation

5. **Error Handling**
   - Invalid coordinates and ranges
   - File operation errors
   - Image processing errors

### Test Patterns

```swift
// Basic functionality tests
func testCreateWorkbook() {
    let workbook = XLKit.createWorkbook()
    XCTAssertNotNil(workbook)
    XCTAssertEqual(workbook.getSheets().count, 0)
}

// API usage tests
func testSetCell() {
    var workbook = XLKit.createWorkbook()
    let sheet = workbook.addSheet(name: "TestSheet")
    
    sheet.setCell("A1", string: "Hello World")
    XCTAssertEqual(sheet.getCell("A1"), .string("Hello World"))
}

// Integration tests
func testCSVImport() {
    let csvData = "Name,Age\nJohn,25\nJane,30"
    let workbook = XLKit.createWorkbookFromCSV(csvData: csvData, hasHeader: true)
    
    let sheets = workbook.getSheets()
    XCTAssertEqual(sheets.count, 1)
    XCTAssertEqual(sheets[0].getCell("A2"), .string("John"))
}
```

## Performance Considerations

### Memory Management

- **Large Datasets**: XLKit is optimized for handling large datasets efficiently
- **Cell Storage**: Uses dictionaries for O(1) cell access
- **Image Handling**: Processes images on-demand to minimize memory usage

### Async Operations

- **File I/O**: Use async/await for file operations to avoid blocking
- **Concurrent Access**: Workbook and Sheet classes are Sendable for thread safety
- **Background Processing**: Image processing and XLSX generation can run in background

### Optimization Tips

```swift
// Use batch operations for multiple cells
sheet.setRange("A1:A1000", value: .string("Batch"))

// Use async operations for file I/O
try await XLKit.saveWorkbook(workbook, to: url)

// Clear unused data
sheet.clear()
```

## Code Style and Standards

### Swift 6.0 Compliance

- **Sendable Conformance**: All public types conform to Sendable where appropriate
- **Modern Swift**: Uses latest Swift features and idioms
- **Type Safety**: Strong typing throughout the codebase

### Documentation Standards

- **Doc Comments**: All public APIs have comprehensive documentation
- **Examples**: Code examples in documentation
- **Parameter Documentation**: All parameters are documented

### Formatting Standards

- **4-space Indentation**: Consistent indentation throughout
- **Trailing Commas**: Used for better git diffs
- **Grouped Imports**: Imports are organized and grouped
- **Modern Idioms**: Uses latest Swift best practices

## Development Guidelines

### Adding New Features

1. **Module Placement**: Place new functionality in appropriate modules
2. **API Design**: Follow existing API patterns and conventions
3. **Testing**: Add comprehensive tests for new features
4. **Documentation**: Update documentation and add examples

### Code Organization

```swift
// File header
//
//  Filename.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

// Imports
import Foundation
@preconcurrency import XLKitCore

// MARK: - Section comments
// Implementation
```

### Error Handling Patterns

```swift
// Use specific error types
guard let coordinate = CellCoordinate(excelAddress: address) else {
    throw XLKitError.invalidCoordinate(address)
}

// Provide meaningful error messages
throw XLKitError.fileWriteError("Failed to write to \(url.path)")
```

### Testing Patterns

```swift
// Test basic functionality
func testFeatureName() {
    // Arrange
    let workbook = XLKit.createWorkbook()
    
    // Act
    let result = workbook.someOperation()
    
    // Assert
    XCTAssertEqual(result, expectedValue)
}

// Test error conditions
func testFeatureNameWithInvalidInput() {
    // Arrange
    let invalidInput = "invalid"
    
    // Act & Assert
    XCTAssertThrowsError(try someOperation(invalidInput)) { error in
        XCTAssertEqual(error as? XLKitError, .expectedError)
    }
}
```

## Build and CI

### Package.swift Configuration

```swift
// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "XLKit",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "XLKit", targets: ["XLKit"]),
        .library(name: "XLKitCore", targets: ["XLKitCore"]),
        .library(name: "XLKitFormatters", targets: ["XLKitFormatters"]),
        .library(name: "XLKitImages", targets: ["XLKitImages"]),
        .library(name: "XLKitXLSX", targets: ["XLKitXLSX"])
    ],
    targets: [
        .target(name: "XLKit", dependencies: ["XLKitCore", "XLKitFormatters", "XLKitImages", "XLKitXLSX"]),
        .target(name: "XLKitCore"),
        .target(name: "XLKitFormatters", dependencies: ["XLKitCore"]),
        .target(name: "XLKitImages", dependencies: ["XLKitCore"]),
        .target(name: "XLKitXLSX", dependencies: ["XLKitCore"]),
        .testTarget(name: "XLKitTests", dependencies: ["XLKit"])
    ]
)
```

### CI Requirements

- **macOS Only**: CI runs only on macOS (no cross-platform support)
- **Swift 6.0**: Uses latest Swift version
- **Test Coverage**: All tests must pass
- **Code Formatting**: Code must be properly formatted
- **Documentation**: All public APIs must be documented

## Future Development

### Potential Enhancements

1. **Cell Formatting**: Enhanced formatting options (colors, fonts, borders)
2. **Charts and Graphs**: Support for Excel charts
3. **Formulas**: Enhanced formula support
4. **Conditional Formatting**: Support for conditional cell formatting
5. **Data Validation**: Cell data validation rules
6. **Pivot Tables**: Support for pivot table creation
7. **Macros**: VBA macro support
8. **Password Protection**: Workbook and sheet protection

### Architecture Considerations

- **Modularity**: Maintain clean module boundaries
- **Extensibility**: Design APIs for future enhancements
- **Performance**: Consider performance implications of new features
- **Compatibility**: Maintain backward compatibility

### Integration Points

- **File Formats**: Consider support for other Excel formats (.xls, .xlsm)
- **Platform Support**: Evaluate iOS/Linux support if needed
- **Third-party Libraries**: Consider integration with other Swift libraries
- **Cloud Storage**: Support for cloud storage providers

This comprehensive guide provides all the information an AI agent needs to understand, maintain, and extend the XLKit library effectively.

## Maintenance & Updates

### Documentation Maintenance Requirements

**IMPORTANT**: This AGENT.MD file must be updated whenever:
- New features are added to the library
- Significant architectural changes are made
- New modules or major components are introduced
- API changes or breaking changes occur
- New testing patterns or requirements are established
- Performance optimizations or security improvements are implemented
- Platform requirements or dependencies change

### Update Checklist for AI Agents

When making significant changes to the codebase, ensure this AGENT.MD file is updated to include:

- **New Features**: Document all new functionality, APIs, and capabilities
- **Architecture Changes**: Update module structure and dependencies
- **API Changes**: Document new methods, parameters, and usage patterns
- **Testing Updates**: Include new test patterns and coverage requirements
- **Performance Notes**: Document performance considerations and optimizations
- **Error Handling**: Update error types and handling patterns
- **Examples**: Provide code examples for new features
- **Integration Notes**: Document any new integration requirements

### Version Tracking

Keep track of major updates to this documentation:
- Document the date of significant updates
- Note the version of XLKit when changes were made
- Reference specific commits or issues that prompted updates

### Cross-Reference with .cursorrules

Ensure consistency between AGENT.MD and .cursorrules files:
- Architecture descriptions match between files
- Coding standards are consistent
- Testing requirements are aligned
- Platform constraints are synchronized
- File organization rules are consistent

This ensures that both documentation files remain accurate and useful for AI agents working with the codebase.

## Architecture Update (2025-07-07)
- The XLSX engine now generates all required files for Excel compliance, including docProps, theme, styles, sharedStrings, and all relationship files.
- Worksheet, styles, and workbook XML are generated in a single-line, Excel-compliant format.
- The test runner (`XLKitTestRunner`) now includes automated validation using [CoreXLSX](https://github.com/CoreOffice/CoreXLSX) to ensure every generated file is fully compliant and readable by Excel and third-party tools.

## Compliance & Validation
- Every generated Excel file is validated for structure and content using CoreXLSX.
- Validation checks include: workbook, worksheet, shared strings, styles, and row/cell integrity.
- The test runner will fail if the generated file is not fully compliant.

## Developer Workflow
- To add new Excel features, update the engine and add/extend validation in the test runner.
- Use the test runner to ensure all changes remain Excel-compliant.
- Manual unpacking and inspection of .xlsx files is no longer required for compliance.

## Recent Improvements
- Added all required OpenXML files and relationships for Excel.
- Automated validation with CoreXLSX.
- Removed legacy debug folders from workflow.

## Image Embedding Implementation (2025-01-27)

### Overview
XLKit now supports advanced image embedding with automatic sizing and aspect ratio preservation, matching the quality of professional Excel exports from tools like Notion.

### Key Features

**Aspect Ratio Preservation**
- Images maintain their original proportions regardless of cell dimensions
- Prevents stretching, squashing, or distortion
- Uses precise Excel formulas for sizing calculations

**Automatic Cell Sizing**
- Row height automatically adjusts to match image height
- Column width automatically adjusts to match image width
- Cells perfectly fit the embedded images

**Precise Positioning**
- Images are positioned exactly at cell boundaries
- No offsets or padding that could cause misalignment
- Perfect alignment with cell grid

**Excel Compliance**
- Uses standard Excel formulas for sizing calculations
- Generates proper drawing XML with correct coordinates
- Maintains compatibility with all Excel versions

### Implementation Details

**API Methods**
```swift
// Sheet extension method
try sheet.embedImageAutoSized(
    imageData: Data,
    at: String,
    workbook: Workbook
) throws

// XLKit convenience method
try XLKit.embedImageAutoSized(
    in: Sheet,
    imageData: Data,
    at: String,
    workbook: Workbook
) throws
```

**Workbook Registration**
- Images are registered with both sheet and workbook
- Ensures proper tracking and counting
- Prevents issues with image enumeration

**Drawing XML Generation**
- Generates precise drawing XML with correct coordinates
- Uses Excel's standard positioning system
- Includes proper relationship references

**Cell Sizing Formulas**
- Column width: `(imageWidth * 7.2) / 96` (Excel's pixel-to-point conversion)
- Row height: `imageHeight * 0.75` (Excel's point-to-pixel conversion)
- Maintains perfect aspect ratio preservation

### Lessons Learned from Notion Export

**Image Sizing**
- Notion uses precise pixel-to-point conversions
- Aspect ratio must be preserved at all costs
- Cell dimensions must exactly match image dimensions

**Positioning**
- Images must be positioned exactly at cell boundaries
- No offsets or padding allowed
- Perfect alignment with Excel's grid system

**File Structure**
- All required OpenXML files must be present
- Proper relationships between files
- Correct content type declarations

**Validation**
- Automated validation ensures compliance
- CoreXLSX validation catches structural issues
- Manual inspection no longer required

### Test Implementation

**ImageEmbedGenerators.swift**
- Dedicated test generator for image embedding
- Embeds images from CSV "Image Filename" column
- Creates new "Image" column for embedded images
- Outputs to `Embed-Test-Embed.xlsx`

**Validation Process**
- Generates Excel file with embedded images
- Validates file structure with CoreXLSX
- Confirms image count and positioning
- Ensures aspect ratio preservation

### Usage Examples

**Basic Image Embedding**
```swift
let workbook = XLKit.createWorkbook()
let sheet = workbook.addSheet(name: "Images")

// Embed image with automatic sizing
try sheet.embedImageAutoSized(
    imageData: imageData,
    at: "B2",
    workbook: workbook
)
```

**CSV with Image Embedding**
```swift
// Create workbook from CSV
let workbook = XLKit.createWorkbookFromCSV(csvData: csvData, hasHeader: true)
let sheet = workbook.getSheets()[0]

// Embed images from filename column
for row in 2...sheet.getUsedCells().count {
    let filename = sheet.getCell("B\(row)").stringValue
    if let imageData = loadImageData(filename: filename) {
        try sheet.embedImageAutoSized(
            imageData: imageData,
            at: "C\(row)", // New Image column
            workbook: workbook
        )
    }
}
```

### Performance Considerations

**Memory Management**
- Images are processed on-demand
- Efficient data handling for large images
- Proper cleanup after processing

**File Size**
- Optimized image compression
- Efficient XML generation
- Minimal overhead for embedded images

**Validation Overhead**
- CoreXLSX validation adds minimal overhead
- Ensures file quality and compliance
- Worthwhile trade-off for reliability

### Future Enhancements

**Planned Features**
- Support for image resizing options
- Multiple image formats (SVG, WebP)
- Image compression options
- Batch image processing

**Optimization Opportunities**
- Parallel image processing
- Caching for repeated images
- Progressive image loading
- Memory usage optimization 