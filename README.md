<p align="center">
  <a href="https://github.com/TheAcharya/XLKit"><img src="Assets/XLKit_Icon.png" height="200">
  <h1 align="center">XLKit</h1>
</p>


<p align="center"><a href="https://github.com/TheAcharya/XLKit/blob/main/LICENSE"><img src="http://img.shields.io/badge/license-MIT-lightgrey.svg?style=flat" alt="license"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit"><img src="https://img.shields.io/badge/platform-macOS%20%7C%20iOS-lightgrey.svg?style=flat" alt="platform"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit/actions/workflows/build.yml"><img src="https://github.com/TheAcharya/XLKit/actions/workflows/build.yml/badge.svg" alt="build"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit/actions/workflows/codeql.yml"><img src="https://github.com/TheAcharya/XLKit/actions/workflows/codeql.yml/badge.svg" alt="codeql"/></a></p>

A modern Swift library for creating and manipulating Excel (.xlsx) files on macOS and iOS. XLKit provides a fluent, chainable API that makes Excel file generation effortless while supporting advanced features like image embedding, CSV/TSV import/export, cell formatting, and both synchronous and asynchronous operations. Built with Swift 6.0 and targeting macOS 12+ and iOS 15+, it offers type-safe operations, comprehensive error handling, and security features. iOS support is available and tested in CI/CD.

Purpose-built for [MarkersExtractor](https://github.com/TheAcharya/MarkersExtractor) - a tool for extracting markers from Final Cut Pro FCPXML files and generating comprehensive Excel reports with embedded images, CSV/TSV manifests, and structured data exports. Perfect for professional video editing workflows requiring pixel-perfect image embedding with all video and cinema aspect ratios.

Includes comprehensive security features for production use: rate limiting, security logging, file quarantine, input validation, and optional checksum verification to protect against vulnerabilities and supply chain attacks.

This codebase is developed using AI agents.

> [!IMPORTANT]
> XLKit has yet to be extensively tested in production environments, real-world workflows, or application integration. This library serves as a modernised foundation for AI-assisted development and experimentation with Excel file generation, and image embedding. While comprehensive testing has been performed, production deployment should be approached with appropriate caution and validation. Additionally, this project would not be actively maintained, so please consider this when planning long-term integrations.

## Table of Contents

- [Features](#features)
- [Recent Improvements](#recent-improvements)
- [Requirements](#requirements)
- [Installing](#installing)
- [Quick Start](#quick-start)
- [Documentation](#documentation)
- [Credits](#credits)
- [License](#license)
- [Reporting Bugs](#reporting-bugs)
- [Contribution](#contribution)

## Features

- Effortless API: Fluent, chainable, and bulk data helpers
- Perfect Image Embedding: Pixel-perfect image embedding with automatic aspect ratio preservation
- Professional Video Support: All 17 video and cinema aspect ratios with zero distortion
- Auto Cell Sizing: Automatic column width and row height calculation to perfectly fit images
- Cell Formatting: Comprehensive formatting including font colours, backgrounds, borders, and text alignment (all 5 alignment options each for horizontal and vertical) with proper XML generation
- Border Support: Full border functionality with thin, medium, and thick styles, custom colors, and combined formatting
- Merge Support: Cell merging functionality with complex range support and proper Excel compliance
- CSV/TSV Import/Export: Built-in support for importing and exporting CSV/TSV data (powered by swift-textfile-tools)
- Async & Sync Operations: Save workbooks with one line (async or sync)
- Type-Safe: Strong enums and structs for all data types
- Excel Compliance: Full OpenXML compliance with CoreXLSX validation
- Minimal Dependencies: CoreXLSX, ZIPFoundation, swift-textfile-tools - all reputable, open-source libraries
- Comprehensive Testing: 59 tests with 100% API coverage, including all text alignment options (horizontal, vertical, combined), number formatting, border and merge functionality, column ordering validation, CSV/TSV edge cases, and automated validation
- Security Features: Comprehensive security features for production use

## Requirements

- macOS: 12.0+
- iOS: 15.0+ (available but not tested)
- Swift: 6.0+

## Installing

### Swift Package Manager

XLKit is available through Swift Package Manager. Add it to your project dependencies:

#### Xcode
1. In Xcode, go to **File** → **Add Package Dependencies**
2. Enter the repository URL: `https://github.com/TheAcharya/XLKit.git`
3. Click **Add Package**
4. Select the XLKit product and click **Add Package**

#### Package.swift
Add XLKit to your `Package.swift` dependencies:

```swift
dependencies: [
    .package(url: "https://github.com/TheAcharya/XLKit.git", from: "1.0.12")
]
```

#### Command Line
```bash
swift package add https://github.com/TheAcharya/XLKit.git
```

### Import

Once installed, import XLKit in your Swift files:

```swift
import XLKit
```

### Verify Installation

Test that XLKit is working correctly:

```swift
import XLKit

// Create a simple workbook
let workbook = Workbook()
let sheet = workbook.addSheet(name: "Test")
sheet.setCell("A1", value: .string("Hello, XLKit!"))

// Save to verify everything works
let safeURL = CoreUtils.safeFileURL(for: "test.xlsx")
try workbook.save(to: safeURL)
print("XLKit is working correctly!")
```

## Quick Start

```swift
import XLKit

// 1. Create a workbook and sheet
let workbook = Workbook()
let sheet = workbook.addSheet(name: "Employees")

// 2. Add headers and data (fluent, chainable)
sheet
    .setRow(1, strings: ["Name", "Photo", "Age"])
    .setRow(2, strings: ["Alice", "", "30"])
    .setRow(3, strings: ["Bob", "", "28"])

// 3. Add formatting with font colours
sheet.setCell("A1", string: "Name", format: CellFormat.header())
sheet.setCell("B1", string: "Photo", format: CellFormat.header())
sheet.setCell("C1", string: "Age", format: CellFormat.text(color: "#FF0000"))

// 4. Add a GIF image to a cell with perfect aspect ratio preservation
let gifData = try Data(contentsOf: URL(fileURLWithPath: "alice.gif"))
try await sheet.embedImageAutoSized(gifData, at: "B2", of: workbook)

// 5. Save the workbook (sync or async)
try await workbook.save(to: URL(fileURLWithPath: "employees.xlsx"))
// or
// try workbook.save(to: url)

// iOS Note: For iOS apps, use CoreUtils.safeFileURL() to get a safe file path:
// let safeURL = CoreUtils.safeFileURL(for: "employees.xlsx")
// try await workbook.save(to: safeURL)
```

## Documentation

Full manual and reference documentation has been moved to the **Documentation** folder:

- **[Documentation/README.md](Documentation/README.md)** – Overview of documentation contents
- **[Documentation/Manual.md](Documentation/Manual.md)** – Complete user manual, including:
  - Security features (SecurityManager, rate limiting, logging, quarantine, checksums)
  - Performance considerations and file format support
  - iOS support and file system considerations
  - Core concepts (Workbook, Sheet, cells, coordinates)
  - Basic and advanced usage with code examples
  - CSV/TSV import and export
  - Image support and aspect ratio preservation
  - Number formats, text alignment, and cell formatting
  - Error handling
  - Testing, validation, and code style

## Credits

Created by [Vigneswaran Rajkumar](https://bsky.app/profile/vigneswaranrajkumar.com)

## License

Licensed under the MIT license. See [LICENSE](https://github.com/TheAcharya/XLKit/blob/main/LICENSE) for details.

## Reporting Bugs

For bug reports, feature requests and suggestions you can create a new [issue](https://github.com/TheAcharya/XLKit/issues) to discuss.

## Contribution

Community contributions are welcome and appreciated. Developers are encouraged to fork the repository and submit pull requests to enhance functionality or introduce thoughtful improvements. However, a key requirement is that nothing should break—all existing features and behaviours and logic must remain fully functional and unchanged. Once reviewed and approved, updates will be merged into the main branch.
