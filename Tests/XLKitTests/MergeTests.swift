//
//  MergeTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct MergeTests {
    
    @Test func testMergeCells() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.mergeCells("A1:C1")
        
        let mergedRanges = sheet.getMergedRanges()
        #expect(mergedRanges.count == 1)
        #expect(mergedRanges[0].excelRange == "A1:C1")
    }
    
    @Test func testMergedCellsActuallyWork() throws {
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
        #expect(mergedRanges.count == 2)
        
        // Check first merge (A1:C1)
        #expect(mergedRanges.contains { $0.excelRange == "A1:C1" })
        let firstMerge = mergedRanges.first { $0.excelRange == "A1:C1" }
        #expect(firstMerge != nil)
        #expect(firstMerge?.start.excelAddress == "A1")
        #expect(firstMerge?.end.excelAddress == "C1")
        
        // Check second merge (A2:B2)
        #expect(mergedRanges.contains { $0.excelRange == "A2:B2" })
        let secondMerge = mergedRanges.first { $0.excelRange == "A2:B2" }
        #expect(secondMerge != nil)
        #expect(secondMerge?.start.excelAddress == "A2")
        #expect(secondMerge?.end.excelAddress == "B2")
        
        // Verify cell values are preserved
        #expect(sheet.getCell("A1") == .string("Merged A1:C1"))
        #expect(sheet.getCell("A2") == .string("Merged A2:B2"))
        #expect(sheet.getCell("C2") == .string("Single Cell"))
        
        // Test Excel file generation with merged cells
        let tempURL = XLKitTestSupport.makeTempWorkbookURL(prefix: "merge_test")
        
        try workbook.save(to: tempURL)
        #expect(FileManager.default.fileExists(atPath: tempURL.path))
        try FileManager.default.removeItem(at: tempURL)
    }
    
    @Test func testComplexMerging() {
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
        #expect(mergedRanges.count == mergeRanges.count)
        
        for range in mergeRanges {
            #expect(mergedRanges.contains { $0.excelRange == range })
        }
        
        // Test specific merge validations
        let horizontalMerge = mergedRanges.first { $0.excelRange == "A1:B1" }
        #expect(horizontalMerge != nil)
        #expect(horizontalMerge?.start.row == 1)
        #expect(horizontalMerge?.start.column == 1)
        #expect(horizontalMerge?.end.row == 1)
        #expect(horizontalMerge?.end.column == 2)
        
        let verticalMerge = mergedRanges.first { $0.excelRange == "A2:A4" }
        #expect(verticalMerge != nil)
        #expect(verticalMerge?.start.row == 2)
        #expect(verticalMerge?.start.column == 1)
        #expect(verticalMerge?.end.row == 4)
        #expect(verticalMerge?.end.column == 1)
        
        let largeMerge = mergedRanges.first { $0.excelRange == "C1:D3" }
        #expect(largeMerge != nil)
        #expect(largeMerge?.start.row == 1)
        #expect(largeMerge?.start.column == 3)
        #expect(largeMerge?.end.row == 3)
        #expect(largeMerge?.end.column == 4)
    }
    
    @Test func testComplexBorderAndMergeCombination() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border and Merge Test")
        
        // Test the exact scenario from the user's report
        var borderedFormat = CellFormat.bordered()
        borderedFormat.fontSize = XLKitTestSupport.standardFontSize
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
        #expect(mergedRanges.count == 2)
        #expect(mergedRanges.contains { $0.excelRange == "A1:B1" })
        #expect(mergedRanges.contains { $0.excelRange == "A2:B2" })
        
        // Verify border formatting
        let borderedCell = sheet.getCellWithFormat("A1")
        #expect(borderedCell != nil)
        #expect(borderedCell?.format?.fontSize == XLKitTestSupport.standardFontSize)
        #expect(borderedCell?.format?.fontWeight == .bold)
        #expect(borderedCell?.format?.horizontalAlignment == .center)
        #expect(borderedCell?.format?.verticalAlignment == .center)
        #expect(borderedCell?.format?.fontName == "Calibri")
        #expect(borderedCell?.format?.borderTop == .thin)
        #expect(borderedCell?.format?.borderBottom == .thin)
        #expect(borderedCell?.format?.borderLeft == .thin)
        #expect(borderedCell?.format?.borderRight == .thin)
        #expect(borderedCell?.format?.borderColor == "#000000")
        
        // Verify column widths
        #expect(sheet.getColumnWidth(1) == 100.0)
        #expect(sheet.getColumnWidth(2) == 150.0)
        
        // Test Excel file generation
        let tempURL = XLKitTestSupport.makeTempWorkbookURL(prefix: "border_merge_test")
        
        try workbook.save(to: tempURL)
        #expect(FileManager.default.fileExists(atPath: tempURL.path))
        try FileManager.default.removeItem(at: tempURL)
    }
}
