# Changelog

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
