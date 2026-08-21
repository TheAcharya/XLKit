//
//  CellFormatInterningTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct CellFormatInterningTests {
    
    @Test func testSetAndGetCellFormatRoundTrip() {
        let sheet = Workbook().addSheet(name: "Format Round Trip")
        let format = CellFormat.header()
        
        sheet.setCellFormat(format, at: "a1")
        
        #expect(sheet.getCellFormat("A1") == format)
        #expect(sheet.getCellFormat("a1") == format)
    }
    
    @Test func testSharedFormatAcrossManyCells() {
        let sheet = Workbook().addSheet(name: "Shared Format")
        let format = CellFormat.header()
        
        for row in 1...20 {
            sheet.setCell("A\(row)", string: "Row \(row)", format: format)
        }
        
        for row in 1...20 {
            #expect(sheet.getCellFormat("A\(row)") == format)
        }
        #expect(sheet.cellFormats.count == 20)
    }
    
    @Test func testRemoveCellFormatLeavesValue() {
        let sheet = Workbook().addSheet(name: "Remove Format")
        sheet.setCell("A1", string: "Hello", format: CellFormat.header())
        
        sheet.removeCellFormat(at: "a1")
        
        #expect(sheet.getCell("A1") == .string("Hello"))
        #expect(sheet.getCellFormat("A1") == nil)
        #expect(sheet.cellFormats.isEmpty)
    }
    
    @Test func testCellFormatsGetterRebuildsDictionary() {
        let sheet = Workbook().addSheet(name: "Formats Getter")
        let header = CellFormat.header()
        
        sheet.setCell("A1", string: "Formatted", format: header)
        sheet.setCell("B1", string: "Plain")
        
        #expect(sheet.cellFormats.count == 1)
        #expect(sheet.cellFormats["A1"] == header)
        #expect(sheet.cellFormats["B1"] == nil)
    }
    
    @Test func testCellFormatsSetterReinterns() {
        let sheet = Workbook().addSheet(name: "Formats Setter")
        sheet.setCell("A1", string: "Old", format: CellFormat.header())
        sheet.setCell("B1", string: "Keep")
        
        let currency = CellFormat.currency()
        let date = CellFormat.date()
        sheet.cellFormats = [
            "A1": currency,
            "c1": date,
        ]
        
        #expect(sheet.getCellFormat("A1") == currency)
        #expect(sheet.getCellFormat("C1") == date)
        #expect(sheet.getCellFormat("B1") == nil)
        #expect(sheet.cellFormats.count == 2)
        #expect(sheet.getCell("A1") == .string("Old"))
        #expect(sheet.getCell("B1") == .string("Keep"))
    }
    
    @Test func testCellFormatEqualityAndHashable() {
        let first = CellFormat.header()
        let second = CellFormat.header()
        let different = CellFormat.currency()
        
        #expect(first == second)
        #expect(first != different)
        
        var uniqueFormats: Set<CellFormat> = []
        uniqueFormats.insert(first)
        uniqueFormats.insert(second)
        uniqueFormats.insert(different)
        
        #expect(uniqueFormats.count == 2)
    }
    
    @Test func testSavingSheetWithSharedFormatsSucceeds() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Shared Format Save")
        let format = CellFormat.header()
        
        for row in 1...50 {
            sheet.setCell("A\(row)", string: "Row \(row)", format: format)
        }
        
        let url = XLKitTestSupport.makeTempWorkbookURL(prefix: "shared_format_save")
        defer { XLKitTestSupport.cleanupTempFile(at: url) }
        
        try workbook.save(to: url)
        
        #expect(FileManager.default.fileExists(atPath: url.path))
        let fileData = try Data(contentsOf: url)
        #expect(!(fileData.isEmpty))
        #expect(sheet.getCellFormat("A1") == format)
        #expect(sheet.getCellFormat("A50") == format)
    }
}
