//
//  TestGeneratorTemplate.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//
//  Template for creating new Excel test generators
//

import Foundation
import XLKit

/// Template function for creating new Excel test generators
/// Copy this function and modify for your specific test case
func templateTestGenerator() throws {
    // MARK: - Configuration
    
    let testName = "Template Test"
    let inputDataPath = "Test-Data/your-input-file.csv"
    let outputExcelFile = "Test-Workflows/Template-Test.xlsx"
    
    print("[INFO] Starting \(testName)...")
    
    // MARK: - Test Logic
    
    print("[INFO] Creating Excel workbook...")
    let workbook = XLKit.createWorkbook()
    let sheet = workbook.addSheet(name: "Template Sheet")
    
    print("[INFO] Adding test data...")
    
    // Header with formatting
    sheet.setCell(row: 1, column: 1, value: .string("Template Test"))
        .setCellFormat(row: 1, column: 1, format: CellFormat(fontWeight: .bold, backgroundColor: "#E0E0E0"))
    
    // Column headers
    sheet.setCell(row: 2, column: 1, value: .string("Column 1"))
    sheet.setCell(row: 2, column: 2, value: .string("Column 2"))
    sheet.setCell(row: 2, column: 3, value: .string("Column 3"))
    
    // Different data types
    sheet.setCell(row: 3, column: 1, value: .number(123.45))
    sheet.setCell(row: 3, column: 2, value: .integer(42))
    sheet.setCell(row: 3, column: 3, value: .boolean(true))
    sheet.setCell(row: 3, column: 4, value: .date(Date()))
    sheet.setCell(row: 3, column: 5, value: .formula("=A3+B3"))
    
    // Image embedding (uncomment if needed)
    // let imageData = try Data(contentsOf: URL(fileURLWithPath: "path/to/image.png"))
    // try XLKit.embedImageAutoSized(imageData, at: "E1", in: sheet, of: workbook)
    
    // Column and row sizing
    sheet.setColumnWidth(1, width: 15.0)
    sheet.setRowHeight(1, height: 20.0)
    
    // Add your specific test logic here
    
    // MARK: - Output
    
    let outputURL = URL(fileURLWithPath: outputExcelFile)
    let outputDir = outputURL.deletingLastPathComponent()
    if !FileManager.default.fileExists(atPath: outputDir.path) {
        print("[INFO] Creating output directory: \(outputDir.path)")
        try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
    }
    
    print("[INFO] Saving Excel file...")
    try XLKit.saveWorkbook(workbook, to: outputURL)
    print("[SUCCESS] Excel file created: \(outputExcelFile)")
    
    // MARK: - Validation
    
    print("[VALIDATION] Validating Excel file: \(outputExcelFile)")
    // Add validation logic here if needed
    print("[VALIDATION] All validation tests passed!")
}

// MARK: - Usage Instructions

/*

TO USE THIS TEMPLATE:

1. Copy template:
   cp Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift Sources/XLKitTestRunner/YourTestName.swift

2. Rename function:
   func yourTestName() throws {

3. Update configuration:
   - testName: Describe your test
   - outputExcelFile: Your desired filename
   - Add specific configuration variables

4. Implement test logic:
   - Replace example code with your requirements
   - Use demonstrated patterns for consistency
   - Add appropriate logging

5. Add to main.swift:
   case "your-test-name":
       do {
           try yourTestName()
       } catch {
           print("[ERROR] Test failed: \(error)")
           exit(1)
       }

6. Update help text in main.swift

7. Create GitHub Actions workflow if needed

FEATURES TO CONSIDER:

Security: Rate limiting, file quarantine, checksums, input validation
Error Handling: All functions throw errors, use do-catch blocks
Image Embedding: embedImageAutoSized for perfect aspect ratio preservation
Cell Formatting: CellFormat for styling (bold, colors, alignment)
Data Types: string, number, integer, boolean, date, formula
Sizing: Column width, row height, image sizing with aspect ratio preservation
Validation: CoreXLSX validation for Excel compliance

NAMING CONVENTIONS:

Function Names (camelCase):
- generateExcelWithImages()
- testCSVImport()
- testCellFormatting()

Test Types (kebab-case):
- csv-import
- image-embedding
- cell-formatting

Output Files:
- Use descriptive names: "Your-Test-Name.xlsx"
- Place in Test-Workflows/ directory

LOGGING:
- [INFO] for general information
- [ERROR] for error conditions
- [SUCCESS] for successful operations
- [VALIDATION] for validation results

*/ 