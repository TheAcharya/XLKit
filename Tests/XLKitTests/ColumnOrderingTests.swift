//
//  ColumnOrderingTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct ColumnOrderingTests {
    
    @Test func testColumnOrderingBeyondZ() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Column Order Test")
        
        // Test the exact scenario from the bug report: 27 columns (A through AA)
        for i in 1...27 {
            sheet.setCell(row: 1, column: i, cell: Cell.string(String(i)))
        }
        
        // Verify cells are set correctly
        #expect(sheet.getCell("A1") == .string("1"))
        #expect(sheet.getCell("Z1") == .string("26"))
        #expect(sheet.getCell("AA1") == .string("27"))
        
        // Test Excel file generation with proper column ordering
        let tempURL = XLKitTestSupport.makeTempWorkbookURL(prefix: "column_order_test")
        defer { XLKitTestSupport.cleanupTempFile(at: tempURL) }
        
        try workbook.save(to: tempURL)
        #expect(FileManager.default.fileExists(atPath: tempURL.path))
        
        let fileData = try Data(contentsOf: tempURL)
        #expect(!(fileData.isEmpty))
    }
    
    @Test func testColumnOrderingWithGaps() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Column Order Gap Test")
        
        // Test with gaps in column usage to ensure proper ordering
        let testColumns = ["A", "B", "Z", "AA", "AB", "AZ", "BA", "BB"]
        for (index, column) in testColumns.enumerated() {
            sheet.setCell("\(column)1", string: "Value \(index + 1)")
        }
        
        // Verify cells are set correctly
        #expect(sheet.getCell("A1") == .string("Value 1"))
        #expect(sheet.getCell("Z1") == .string("Value 3"))
        #expect(sheet.getCell("AA1") == .string("Value 4"))
        #expect(sheet.getCell("BA1") == .string("Value 7"))
        
        // Test Excel file generation
        let tempURL = XLKitTestSupport.makeTempWorkbookURL(prefix: "column_order_gap_test")
        defer { XLKitTestSupport.cleanupTempFile(at: tempURL) }
        
        try workbook.save(to: tempURL)
        #expect(FileManager.default.fileExists(atPath: tempURL.path))
        
        let fileData = try Data(contentsOf: tempURL)
        #expect(!(fileData.isEmpty))
    }
}
