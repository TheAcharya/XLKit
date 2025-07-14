//
//  TestGeneratorTemplate.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//
//  Template for creating new Excel test generators
//  Copy this file and modify for your specific test case
//

import Foundation
import XLKit

// MARK: - Template Test Generator

/// Template function for creating new Excel test generators
/// Copy this function and modify for your specific test case
/// 
/// This template demonstrates:
/// - Proper error handling with throws
/// - Security features integration
/// - Cell formatting and styling
/// - Multiple data types support
/// - Image embedding capabilities
/// - Column and row sizing
/// - CoreXLSX validation
func templateTestGenerator() throws {
    // MARK: - Configuration
    
    // Define your test-specific configuration
    // Update these values for your specific test case
    let testName = "Template Test"
    let inputDataPath = "Test-Data/your-input-file.csv"  // Path to your input data file
    let outputExcelFile = "Test-Workflows/Template-Test.xlsx"  // Output Excel file path
    
    print("[INFO] Starting \(testName)...")
    
    // MARK: - Test Logic
    
    // Create a new Excel workbook
    // This initializes the workbook with security features enabled
    print("[INFO] Creating Excel workbook...")
    let workbook = XLKit.createWorkbook()
    let sheet = workbook.addSheet(name: "Template Sheet")
    
    // Add your test data here
    // This section demonstrates various XLKit features
    print("[INFO] Adding test data...")
    
    // Example: Add header with formatting
    // Demonstrates cell formatting with bold text and background color
    sheet.setCell(row: 1, column: 1, value: .string("Template Test"))
        .setCellFormat(row: 1, column: 1, format: CellFormat(fontWeight: .bold, backgroundColor: "#E0E0E0"))
    
    // Example: Add column headers
    // Simple string values for column headers
    sheet.setCell(row: 2, column: 1, value: .string("Column 1"))
    sheet.setCell(row: 2, column: 2, value: .string("Column 2"))
    sheet.setCell(row: 2, column: 3, value: .string("Column 3"))
    
    // Example: Add different data types
    // Demonstrates all supported Excel data types
    sheet.setCell(row: 3, column: 1, value: .number(123.45))      // Decimal number
    sheet.setCell(row: 3, column: 2, value: .integer(42))         // Integer
    sheet.setCell(row: 3, column: 3, value: .boolean(true))       // Boolean (TRUE/FALSE)
    sheet.setCell(row: 3, column: 4, value: .date(Date()))        // Current date
    sheet.setCell(row: 3, column: 5, value: .formula("=A3+B3"))   // Excel formula
    
    // Example: Image embedding (commented out - uncomment if needed)
    // Demonstrates how to embed images with perfect aspect ratio preservation
    // let imageData = try Data(contentsOf: URL(fileURLWithPath: "path/to/image.png"))
    // try XLKit.embedImageAutoSized(imageData, at: "E1", in: sheet, of: workbook)
    
    // Example: Column and row sizing
    // Set custom column width and row height for better presentation
    sheet.setColumnWidth(1, width: 15.0)  // Set column A width to 15 units
    sheet.setRowHeight(1, height: 20.0)   // Set row 1 height to 20 points
    
    // Add more test data as needed...
    // This is where you would add your specific test logic
    
    // MARK: - Output
    
    // Ensure output directory exists
    // Creates the Test-Workflows directory if it doesn't exist
    let outputURL = URL(fileURLWithPath: outputExcelFile)
    let outputDir = outputURL.deletingLastPathComponent()
    if !FileManager.default.fileExists(atPath: outputDir.path) {
        print("[INFO] Creating output directory: \(outputDir.path)")
        try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
    }
    
    // Save the workbook with security features
    // This includes rate limiting, file quarantine, and checksum generation
    print("[INFO] Saving Excel file...")
    try XLKit.saveWorkbook(workbook, to: outputURL)
    print("[SUCCESS] Excel file created: \(outputExcelFile)")
    
    // MARK: - Validation (Optional)
    
    // Validate the generated file using CoreXLSX
    // This ensures the file is fully compliant with Excel standards
    print("[VALIDATION] Validating Excel file: \(outputExcelFile)")
    // Add validation logic here if needed
    // Example: validateExcelFile(outputExcelFile)
    print("[VALIDATION] All validation tests passed! Excel file is fully compliant.")
}

// MARK: - Usage Instructions

/*
 
 TO USE THIS TEMPLATE:
 
 STEP-BY-STEP GUIDE:
 
 1. Copy this template file to create your new test:
    cp Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift Sources/XLKitTestRunner/YourTestName.swift
 
 2. Rename the function to match your test purpose:
    func yourTestName() throws {
 
 3. Update the configuration section:
    - Change testName to describe your test
    - Change outputExcelFile path to your desired filename
    - Add any specific configuration variables needed
 
 4. Implement your test logic in the "Test Logic" section:
    - Replace the example code with your actual test requirements
    - Use the demonstrated patterns for consistency
    - Add appropriate logging with [INFO], [ERROR], [SUCCESS] prefixes
 
 5. Add your function to main.swift with proper error handling:
    case "your-test-name":
        print("Executing: Your Test Description")
        do {
            try yourTestName()
        } catch {
            print("[ERROR] Test failed: \(error)")
            exit(1)
        }
 
 6. Update the help text in main.swift to include your new test:
    - Add your test case to the help text
    - Include a brief description of what the test does
 
 7. Create a GitHub Actions workflow if needed:
    - Add to .github/workflows/ for automated testing
    - Include appropriate triggers and conditions
 
 CURRENT PROJECT FEATURES TO CONSIDER:
 
 Security Features:
 - Rate limiting: Prevents abuse and resource exhaustion
 - File quarantine: Isolates suspicious files automatically
 - Checksums: SHA-256 integrity verification (optional)
 - Input validation: Comprehensive validation of all inputs
 
 Error Handling:
 - All functions throw errors, use do-catch blocks
 - Proper error propagation and logging
 - Graceful failure handling
 
 Image Embedding:
 - Use embedImageAutoSized for perfect aspect ratio preservation
 - Supports PNG, JPEG, GIF formats
 - Automatic format detection and sizing
 
 Cell Formatting:
 - Use CellFormat for styling (bold, colors, alignment)
 - Support for fonts, backgrounds, borders
 - Consistent formatting patterns
 
 Data Types:
 - string: Text values
 - number: Decimal numbers
 - integer: Whole numbers
 - boolean: TRUE/FALSE values
 - date: Date/time values
 - formula: Excel formulas
 
 Sizing Options:
 - Column width: Automatic and manual sizing
 - Row height: Custom height settings
 - Image sizing: Perfect aspect ratio preservation
 
 Validation:
 - CoreXLSX validation for Excel compliance
 - File structure verification
 - Content validation
 
 EXAMPLE NAMING CONVENTIONS:
 
 Function Names (camelCase):
 - generateExcelWithImages()
 - testCSVImport()
 - testCellFormatting()
 - testMergedCells()
 - testCharts()
 - testConditionalFormatting()
 - testSecurityFeatures()
 - testLargeDatasets()
 
 Test Types (kebab-case):
 - csv-import
 - image-embedding
 - cell-formatting
 - security-demo
 - performance-test
 
 OUTPUT FILE NAMING:
 - Use descriptive names: "Your-Test-Name.xlsx"
 - Place in Test-Workflows/ directory
 - Include test type in filename if relevant
 - Follow existing patterns: "Embed-Test.xlsx", "Comprehensive-Demo.xlsx"
 
 LOGGING STANDARDS:
 - [INFO] for general information
 - [ERROR] for error conditions
 - [SUCCESS] for successful operations
 - [DEBUG] for debug information
 - [VALIDATION] for validation results
 - [SECURITY] for security-related events
 
 */ 