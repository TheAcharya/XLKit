# Changelog

### 1.0.11

**🎉 Released:**
- 8th February 2026

**🔧 Improvements:**
- Fixed documentation inconsistency: replaced all references to `CellFormat.coloredText(color:)` with `CellFormat.text(color:)` in README.md and AGENT.MD to match the current API
- Removed force unwraps in `CoreTypes.swift` for Swift 6 compliance and safer behaviour: `CellCoordinate.init?(excelAddress:)` now uses optional binding for `Range` conversion; `CoreUtils.columnLetter(from:)` uses optional binding for `UnicodeScalar`; `CoreUtils.columnNumber(from:)` safely handles non-ASCII characters instead of crashing
- Removed force unwraps in `XLSXEngine.generateCellXML()`: replaced `format!` and `styleId!` with optional binding and `styleId.map` for style attribute generation
- Updated unit tests to use `FileManager.default.temporaryDirectory` instead of `NSTemporaryDirectory()` for cross-platform consistency and alignment with project guidelines

---

### 1.0.10

**🎉 Released:**
- 18th October 2025

**🔧 Improvements:**
- Fixed critical column ordering bug for sheets with more than 26 columns (A-Z, AA, AB, etc.) (#14)
- Resolved Excel compatibility issue where generated files were rejected or repaired due to invalid column ordering
- Implemented proper numeric column sorting in `XLSXEngine.generateWorksheetXML()` to ensure Excel-compliant column order
- Fixed lexicographic string sorting that caused `AA1` to appear before `B1` in generated XML
- Increased test count from 53 to 55 tests with 100% API coverage including column ordering validation
- Enhanced XLKitTestRunner with column ordering validation for continuous testing 

---

### 1.0.9

**🎉 Released:**
- 25th September 2025

**🔧 Improvements:**
- Fixed text wrapping functionality that was defined in CellFormat but not being used in XLSX generation (#12)
- Added text wrapping support to XLSXEngine alignment XML generation with proper wrapText attribute
- Enhanced formatToKey function to include text wrapping in format key generation for proper format caching
- Increased test count from 51 to 53 tests with 100% API coverage including text wrapping validation
- Implemented proper Excel XML structure with alignment elements containing wrapText attribute

---

### 1.0.8

**🎉 Released:**
- 5th August 2025

**🔧 Improvements:**
- Added comprehensive border and merge functionality with full Excel compliance (#7)
- Implemented border support with thin, medium, and thick styles with custom colours
- Added merged cells support with complex range scenarios and proper XML generation
- Enhanced CellFormat struct with border properties (borderTop, borderBottom, borderLeft, borderRight, borderColor)
- Added bordered() factory method for easy border formatting creation
- Implemented proper border XML generation in styles.xml with dynamic border definitions
- Added mergeCells() functionality with proper range validation and storage
- Enhanced XLSXEngine to generate mergeCells XML section in worksheet files
- Added comprehensive border and merge testing with 6 new test functions
- Increased test count from 45 to 51 tests with 100% API coverage including border and merge validation
- Added border and merge combinations with other formatting options (font, alignment, background)
- Validated border and merge functionality with CoreXLSX to ensure full Excel compliance

---

### 1.0.7

**🎉 Released:**
- 4th August 2025

**🔧 Improvements:**
- Fixed custom number formats being silently ignored by Excel in generated XLSX files (#8)
- Implemented proper `<numFmts>` section in `styles.xml` with dynamic number format ID assignment
- Added `applyNumberFormat="1"` attribute to cell format (`xf`) elements for Excel compliance
- Updated `formatToKey` function to include number format information for unique style generation
- Added comprehensive number formatting tests with 5 new test functions covering currency, percentage, and custom formats
- Increased test count from 40 to 45 tests with 100% API coverage including number formatting validation
- Added `number-formats` test type to XLKitTestRunner for end-to-end number formatting validation
- Enhanced comprehensive demo with number formatting examples for currency and percentage display
- Ensured all generated Excel files correctly display thousands grouping, currency symbols, and custom formats
- Fixed Excel "Format Cells" dialog to properly reflect custom number formats instead of showing "General"
- Validated number formatting fix with CoreXLSX to ensure full Excel compliance and compatibility

---

### 1.0.6

**🎉 Released:**
- 18th July 2025

**🔧 Improvements:**
- Fixed iOS "ZIP creation error" issue where users couldn't save Excel files on iOS (#5)
- Implemented iOS-specific file operations using copy instead of move for better sandbox compatibility
- Added `CoreUtils.safeFileURL()` helper method for iOS-safe file paths
- Updated file path validation to be more permissive on iOS while maintaining security on macOS
- Added iOS compatibility test to XLKitTestRunner for continuous validation

---

### 1.0.5

**🎉 Released:**
- 15th July 2025

**🔧 Improvements:**
- Added comprehensive text alignment testing with 5 new test functions covering all Excel alignment options
- Increased test count from 35 to 40 tests with 100% API coverage including all text alignment options
- Optimised code comments throughout the codebase for better maintainability and readability
- Preserved essential API documentation and implementation details while removing unnecessary verbosity
- Added detailed Text Alignment Support section in README.md with practical examples and usage patterns

---

### 1.0.4

**🎉 Released:**
- 14th July 2025

**🔧 Improvements:**
- Added iOS platform support with iOS 15+ targeting and cross-platform compatibility
- Fixed iOS build error: `'homeDirectoryForCurrentUser' is unavailable in iOS`
- Implemented platform-specific conditionals for file system operations using `#if os(macOS)`
- Updated `allowedDirectories` in CoreTypes.swift to use platform-specific home directory access
- Added iOS job to GitHub Actions workflow for continuous testing on iOS simulators
- Added CodeQL security scanning workflow with Swift 6.0 support
- Fixed comprehensive demo test with duplicate relationship IDs causing Excel file corruption
- Implemented dynamic relationship ID generation to prevent conflicts in `workbook.xml.rels` and `drawing1.xml.rels`
- Ensured all generated Excel files pass `CoreXLSX` validation and open without errors
- Fixed async/await warnings by removing unnecessary await keywords from synchronous operations
- Updated API documentation to show instance methods on Workbook and Sheet classes
- Updated API documentation to use proper Swift naming conventions and parameter patterns
- Fixed concurrency handling in test runner with proper Task and semaphore usage
- Added font colour support with proper XML generation in `XLSXEngine`
- Updated `CellFormat` to properly apply font colours in Excel output
- Added font colour formatting tests to XLKitTests with comprehensive validation
- Enhanced comprehensive demo with font colour examples for all supported colours
- Fixed `Image` column header formatting inconsistency in embed test output
- Ensured all generated Excel files maintain consistent header styling across all columns
- Enhanced codebase with modern Swift 6.0 idioms and improved type safety
- Improved error handling with more specific `XLKitError` types and meaningful error messages
- Enhanced test suite organisation with better test categorisation and coverage documentation
- Resolved test count discrepancies and ensured accurate documentation across all files
- Improved code maintainability with consistent formatting and documentation standards

---

### 1.0.3

**🎉 Released:**
- 12th July 2025

**🔧 Improvements:**
- Identified and resolved scaling inconsistencies between XLKit test and MarkersExtractor integration
- Fixed XLKit test to use default parameters instead of manual overrides for consistent behaviour
- Established consistent API usage pattern: let XLKit handle sizing automatically
- Documented scale parameter options and best practices for image embeddings
- Updated documentation (AGENT.MD and .cursorrules) with scaling API details and integration guidelines

---

### 1.0.2

**🎉 Released:**
- 9th July 2025

**🔧 Improvements:**
- Updated Security defaults

---

### 1.0.1

**🎉 Released:**
- 9th July 2025

**🔧 Improvements:**
- Added SecurityManager with rate limiting, security logging, file quarantine, and checksums
- Integrated security features throughout XLSXEngine, ImageUtils, and XLKit API
- Added comprehensive input validation for all user inputs
- Replaced system zip command with pure Swift ZIP library to eliminate command injection risk
- Fixed Swift 6.0 concurrency issues with @MainActor and throws propagation
- Enhanced XLKitTestRunner with security integration and better error handling
- Updated documentation (AGENT.MD and .cursorrules) with security features
- Fixed test runner hanging issues and build errors
- Expanded from 9 to all 17 professional video and cinema aspect ratios with pixel-perfect preservation
- Added comprehensive security requirements and validation processes to development guidelines

---

### 1.0.0

**🎉 Released:**
- 7th July 2025

This is the first public release of **XLKit**!
