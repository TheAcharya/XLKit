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
func templateTestGenerator() {
    // MARK: - Configuration
    
    // Define your test-specific configuration
    let testName = "Template Test"
    let inputDataPath = "Test-Data/your-input-file.csv"  // Update this path
    let outputExcelFile = "Test-Workflows/Template-Test.xlsx"
    
    print("[INFO] Starting \(testName)...")
    
    // MARK: - Test Logic
    
    // Create workbook
    print("[INFO] Creating Excel workbook...")
    let workbook = XLKit.createWorkbook()
    let sheet = workbook.addSheet(name: "Template Sheet")
    
    // Add your test data here
    print("[INFO] Adding test data...")
    
    // Example: Add some sample data
    sheet.setCell(row: 1, column: 1, value: .string("Template Test"))
    sheet.setCell(row: 2, column: 1, value: .string("Column 1"))
    sheet.setCell(row: 2, column: 2, value: .string("Column 2"))
    sheet.setCell(row: 2, column: 3, value: .string("Column 3"))
    
    // Add more test data as needed...
    
    // MARK: - Output
    
    // Ensure output directory exists
    let outputURL = URL(fileURLWithPath: outputExcelFile)
    let outputDir = outputURL.deletingLastPathComponent()
    if !FileManager.default.fileExists(atPath: outputDir.path) {
        print("[INFO] Creating output directory: \(outputDir.path)")
        do {
            try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
        } catch {
            print("[ERROR] Failed to create output directory: \(error)")
            exit(1)
        }
    }
    
    // Save the workbook
    print("[INFO] Saving Excel file...")
    do {
        try XLKit.saveWorkbook(workbook, to: outputURL)
        print("[SUCCESS] Excel file created: \(outputExcelFile)")
    } catch {
        print("[ERROR] Failed to save Excel file: \(error)")
        exit(1)
    }
}

// MARK: - Usage Instructions

/*
 
 TO USE THIS TEMPLATE:
 
 1. Copy this file to a new file with your test name:
    cp Sources/XLKitTestRunner/Templates/TestGeneratorTemplate.swift Sources/XLKitTestRunner/YourTestName.swift
 
 2. Rename the function:
    func yourTestName() {
 
 3. Update the configuration section:
    - Change testName
    - Change outputExcelFile path
    - Add any specific configuration needed
 
 4. Implement your test logic in the "Test Logic" section
 
 5. Add your function to main.swift:
    case "your-test-name":
        print("Executing: Your Test Description")
        yourTestName()
 
 6. Update the help text in main.swift to include your new test
 
 7. Create a GitHub Actions workflow if needed
 
 EXAMPLE NAMING CONVENTIONS:
 - generateExcelWithImages()
 - testCSVImport()
 - testCellFormatting()
 - testMergedCells()
 - testCharts()
 - testConditionalFormatting()
 
 */ 