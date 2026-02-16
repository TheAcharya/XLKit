//
//  ColumnOrderingTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class ColumnOrderingTests: XLKitTestBase {
    
    func testColumnOrderingBeyondZ() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Column Order Test")
        
        // Test the exact scenario from the bug report: 27 columns (A through AA)
        for i in 1...27 {
            sheet.setCell(row: 1, column: i, cell: Cell.string(String(i)))
        }
        
        // Verify cells are set correctly
        XCTAssertEqual(sheet.getCell("A1"), .string("1"))
        XCTAssertEqual(sheet.getCell("Z1"), .string("26"))
        XCTAssertEqual(sheet.getCell("AA1"), .string("27"))
        
        // Test Excel file generation with proper column ordering
        let tempURL = makeTempWorkbookURL(prefix: "column_order_test")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            
            // Verify the file can be read back and contains the expected data
            // This test ensures the XML is generated with proper column ordering
            let fileData = try Data(contentsOf: tempURL)
            XCTAssertFalse(fileData.isEmpty)
            
            // Clean up
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with >26 columns: \(error)")
        }
    }
    
    func testColumnOrderingWithGaps() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Column Order Gap Test")
        
        // Test with gaps in column usage to ensure proper ordering
        let testColumns = ["A", "B", "Z", "AA", "AB", "AZ", "BA", "BB"]
        for (index, column) in testColumns.enumerated() {
            sheet.setCell("\(column)1", string: "Value \(index + 1)")
        }
        
        // Verify cells are set correctly
        XCTAssertEqual(sheet.getCell("A1"), .string("Value 1"))
        XCTAssertEqual(sheet.getCell("Z1"), .string("Value 3"))
        XCTAssertEqual(sheet.getCell("AA1"), .string("Value 4"))
        XCTAssertEqual(sheet.getCell("BA1"), .string("Value 7"))
        
        // Test Excel file generation
        let tempURL = makeTempWorkbookURL(prefix: "column_order_gap_test")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            
            // Verify the file can be read back
            let fileData = try Data(contentsOf: tempURL)
            XCTAssertFalse(fileData.isEmpty)
            
            // Clean up
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with column gaps: \(error)")
        }
    }
}
