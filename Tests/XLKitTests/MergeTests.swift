//
//  MergeTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class MergeTests: XLKitTestBase {
    
    func testMergeCells() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.mergeCells("A1:C1")
        
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 1)
        XCTAssertEqual(mergedRanges[0].excelRange, "A1:C1")
    }
    
    func testMergedCellsActuallyWork() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Merge Test")
        
        // Test complex merging scenarios
        sheet.setCell("A1", string: "Merged A1:C1")
        sheet.setCell("A2", string: "Merged A2:B2")
        sheet.setCell("C2", string: "Single Cell")
        
        // Perform merges
        sheet.mergeCells("A1:C1")
        sheet.mergeCells("A2:B2")
        
        // Verify merged ranges are stored correctly
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 2)
        
        // Check first merge (A1:C1)
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A1:C1" })
        let firstMerge = mergedRanges.first { $0.excelRange == "A1:C1" }
        XCTAssertNotNil(firstMerge)
        XCTAssertEqual(firstMerge?.start.excelAddress, "A1")
        XCTAssertEqual(firstMerge?.end.excelAddress, "C1")
        
        // Check second merge (A2:B2)
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A2:B2" })
        let secondMerge = mergedRanges.first { $0.excelRange == "A2:B2" }
        XCTAssertNotNil(secondMerge)
        XCTAssertEqual(secondMerge?.start.excelAddress, "A2")
        XCTAssertEqual(secondMerge?.end.excelAddress, "B2")
        
        // Verify cell values are preserved
        XCTAssertEqual(sheet.getCell("A1"), .string("Merged A1:C1"))
        XCTAssertEqual(sheet.getCell("A2"), .string("Merged A2:B2"))
        XCTAssertEqual(sheet.getCell("C2"), .string("Single Cell"))
        
        // Test Excel file generation with merged cells
        let tempURL = makeTempWorkbookURL(prefix: "merge_test")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with merged cells: \(error)")
        }
    }
    
    func testComplexMerging() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Complex Merge Test")
        
        // Test various merge scenarios
        let mergeRanges = [
            "A1:B1",   // Horizontal merge
            "A2:A4",   // Vertical merge
            "C1:D3",   // 2x3 merge
            "F1:F1",   // Single cell (no merge)
            "G1:H2"    // 2x2 merge
        ]
        
        // Set data and perform merges
        for (index, range) in mergeRanges.enumerated() {
            let startCell = range.components(separatedBy: ":")[0]
            sheet.setCell(startCell, string: "Merge \(index + 1): \(range)")
            sheet.mergeCells(range)
        }
        
        // Verify all merges are stored correctly
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, mergeRanges.count)
        
        for range in mergeRanges {
            XCTAssertTrue(mergedRanges.contains { $0.excelRange == range })
        }
        
        // Test specific merge validations
        let horizontalMerge = mergedRanges.first { $0.excelRange == "A1:B1" }
        XCTAssertNotNil(horizontalMerge)
        XCTAssertEqual(horizontalMerge?.start.row, 1)
        XCTAssertEqual(horizontalMerge?.start.column, 1)
        XCTAssertEqual(horizontalMerge?.end.row, 1)
        XCTAssertEqual(horizontalMerge?.end.column, 2)
        
        let verticalMerge = mergedRanges.first { $0.excelRange == "A2:A4" }
        XCTAssertNotNil(verticalMerge)
        XCTAssertEqual(verticalMerge?.start.row, 2)
        XCTAssertEqual(verticalMerge?.start.column, 1)
        XCTAssertEqual(verticalMerge?.end.row, 4)
        XCTAssertEqual(verticalMerge?.end.column, 1)
        
        let largeMerge = mergedRanges.first { $0.excelRange == "C1:D3" }
        XCTAssertNotNil(largeMerge)
        XCTAssertEqual(largeMerge?.start.row, 1)
        XCTAssertEqual(largeMerge?.start.column, 3)
        XCTAssertEqual(largeMerge?.end.row, 3)
        XCTAssertEqual(largeMerge?.end.column, 4)
    }
    
    func testComplexBorderAndMergeCombination() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border and Merge Test")
        
        // Test the exact scenario from the user's report
        var borderedFormat = CellFormat.bordered()
        borderedFormat.fontSize = Self.standardFontSize
        borderedFormat.fontWeight = .bold
        borderedFormat.horizontalAlignment = .center
        borderedFormat.verticalAlignment = .center
        borderedFormat.fontName = "Calibri"
        borderedFormat.borderTop = .thin
        borderedFormat.borderBottom = .thin
        borderedFormat.borderLeft = .thin
        borderedFormat.borderRight = .thin
        borderedFormat.borderColor = "#000000"
        
        // Set cells with borders and formatting
        sheet.setCell("A1", string: "Test1", format: borderedFormat)
        sheet.setCell("A2", string: "Test2")
        
        // Perform merges
        sheet.mergeCells("A1:B1")
        sheet.mergeCells("A2:B2")
        
        // Set column widths
        sheet.setColumnWidth(1, width: 100.0)
        sheet.setColumnWidth(2, width: 150.0)
        
        // Verify everything is stored correctly
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 2)
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A1:B1" })
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A2:B2" })
        
        // Verify border formatting
        let borderedCell = sheet.getCellWithFormat("A1")
        XCTAssertNotNil(borderedCell)
        XCTAssertEqual(borderedCell?.format?.fontSize, Self.standardFontSize)
        XCTAssertEqual(borderedCell?.format?.fontWeight, .bold)
        XCTAssertEqual(borderedCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(borderedCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(borderedCell?.format?.fontName, "Calibri")
        XCTAssertEqual(borderedCell?.format?.borderTop, .thin)
        XCTAssertEqual(borderedCell?.format?.borderBottom, .thin)
        XCTAssertEqual(borderedCell?.format?.borderLeft, .thin)
        XCTAssertEqual(borderedCell?.format?.borderRight, .thin)
        XCTAssertEqual(borderedCell?.format?.borderColor, "#000000")
        
        // Verify column widths
        XCTAssertEqual(sheet.getColumnWidth(1), 100.0)
        XCTAssertEqual(sheet.getColumnWidth(2), 150.0)
        
        // Test Excel file generation
        let tempURL = makeTempWorkbookURL(prefix: "border_merge_test")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with borders and merges: \(error)")
        }
    }
}
