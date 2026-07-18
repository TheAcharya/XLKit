//
//  CSVTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct CSVTests {
    /// Shared CSV fixture used across import/convenience-init tests.
    private static let csvSampleWithHeader = """
    Name,Age,Salary
    Alice,30,50000.5
    Bob,25,45000.75
    """

    /// Shared TSV fixture used across import/convenience-init tests.
    private static let tsvSampleWithHeader = """
    Product\tPrice\tIn Stock
    Apple\t1.99\ttrue
    Banana\t0.99\tfalse
    """
    
    @Test func testWorkbookFromCSV() throws {
        let csvData = Self.csvSampleWithHeader
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        let sheets = workbook.getSheets()
        let sheet = try #require(sheets.first)
        
        #expect(sheet.getCell("A2") == .string("Alice"))
        #expect(sheet.getCell("B2") == .integer(30))
        #expect(sheet.getCell("C2") == .number(50000.5))
        #expect(sheet.getCell("A3") == .string("Bob"))
        #expect(sheet.getCell("B3") == .integer(25))
        #expect(sheet.getCell("C3") == .number(45000.75))
    }
    
    @Test func testWorkbookFromTSV() throws {
        let tsvData = Self.tsvSampleWithHeader
        
        let workbook = Workbook(fromTSV: tsvData, hasHeader: true)
        let sheet = try #require(workbook.getSheets().first)
        
        #expect(sheet.getCell("A2") == .string("Apple"))
        #expect(sheet.getCell("B2") == .number(1.99))
        #expect(sheet.getCell("C2") == .boolean(true))
        #expect(sheet.getCell("A3") == .string("Banana"))
        #expect(sheet.getCell("B3") == .number(0.99))
        #expect(sheet.getCell("C3") == .boolean(false))
    }
    
    @Test func testSheetExportToCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Age", "Salary"])
        sheet.setRow(2, strings: ["Alice", "30", "50000"])
        sheet.setRow(3, strings: ["Bob", "25", "45000"])
        
        let csv = sheet.exportToCSV()
        let expectedCSV = "Name,Age,Salary\nAlice,30,50000\nBob,25,45000"
        
        #expect(csv == expectedCSV)
    }
    
    @Test func testSheetExportToTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Product", "Price", "In Stock"])
        sheet.setRow(2, strings: ["Apple", "1.99", "true"])
        sheet.setRow(3, strings: ["Banana", "0.99", "false"])
        
        let tsv = sheet.exportToTSV()
        let expectedTSV = "Product\tPrice\tIn Stock\nApple\t1.99\ttrue\nBanana\t0.99\tfalse"
        
        #expect(tsv == expectedTSV)
    }
    
    @Test func testWorkbookImportCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let csvData = Self.csvSampleWithHeader
        
        workbook.importCSV(csvData, into: sheet, hasHeader: true)
        
        #expect(sheet.getCell("A2") == .string("Alice"))
        #expect(sheet.getCell("B2") == .integer(30))
        #expect(sheet.getCell("C2") == .number(50000.5))
        #expect(sheet.getCell("A3") == .string("Bob"))
        #expect(sheet.getCell("B3") == .integer(25))
        #expect(sheet.getCell("C3") == .number(45000.75))
    }
    
    @Test func testWorkbookImportTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let tsvData = Self.tsvSampleWithHeader
        
        workbook.importTSV(tsvData, into: sheet, hasHeader: true)
        
        #expect(sheet.getCell("A2") == .string("Apple"))
        #expect(sheet.getCell("B2") == .number(1.99))
        #expect(sheet.getCell("C2") == .boolean(true))
        #expect(sheet.getCell("A3") == .string("Banana"))
        #expect(sheet.getCell("B3") == .number(0.99))
        #expect(sheet.getCell("C3") == .boolean(false))
    }
    
    @Test func testWorkbookExportSheetToCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Age", "Salary"])
        sheet.setRow(2, strings: ["Alice", "30", "50000"])
        sheet.setRow(3, strings: ["Bob", "25", "45000"])
        
        let csv = workbook.exportSheetToCSV(sheet)
        let expectedCSV = "Name,Age,Salary\nAlice,30,50000\nBob,25,45000"
        
        #expect(csv == expectedCSV)
    }
    
    @Test func testWorkbookExportSheetToTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Product", "Price", "In Stock"])
        sheet.setRow(2, strings: ["Apple", "1.99", "true"])
        sheet.setRow(3, strings: ["Banana", "0.99", "false"])
        
        let tsv = workbook.exportSheetToTSV(sheet)
        let expectedTSV = "Product\tPrice\tIn Stock\nApple\t1.99\ttrue\nBanana\t0.99\tfalse"
        
        #expect(tsv == expectedTSV)
    }
    
    @Test func testCSVWithQuotedFields() throws {
        // Test CSV with quoted fields containing commas (swift-textfile handles this)
        let csvData = "Name,Description,Price\nApple,\"Red, delicious apple\",1.99\nBanana,\"Yellow, curved fruit\",0.99"
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        let sheet = try #require(workbook.getSheets().first)
        
        #expect(sheet.getCell("A2") == .string("Apple"))
        #expect(sheet.getCell("B2") == .string("Red, delicious apple"))
        #expect(sheet.getCell("C2") == .number(1.99))
        #expect(sheet.getCell("A3") == .string("Banana"))
        #expect(sheet.getCell("B3") == .string("Yellow, curved fruit"))
        #expect(sheet.getCell("C3") == .number(0.99))
    }
    
    @Test func testCSVWithEscapedQuotes() throws {
        // Test CSV with escaped quotes (double quotes) - swift-textfile handles this
        let csvData = #"""
        Name,Quote
        Alice,"She said ""Hello"""
        Bob,"He said ""Goodbye"""
        """#
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        let sheet = try #require(workbook.getSheets().first)
        
        #expect(sheet.getCell("A2") == .string("Alice"))
        #expect(sheet.getCell("B2") == .string("She said \"Hello\""))
        #expect(sheet.getCell("A3") == .string("Bob"))
        #expect(sheet.getCell("B3") == .string("He said \"Goodbye\""))
    }
    
    @Test func testCSVExportWithSpecialCharacters() throws {
        // Test that CSV export properly quotes fields with special characters
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Description"])
        sheet.setRow(2, strings: ["Apple", "Red, delicious"])
        sheet.setRow(3, strings: ["Quote", "\"Hello\""])
        
        let csv = sheet.exportToCSV()
        
        // Verify the CSV contains quoted fields for special characters
        #expect(csv.contains("\"Red, delicious\""), "CSV should quote fields containing commas")
        
        // Verify round-trip: import the exported CSV (this validates CSV format correctness)
        // The round-trip test ensures escaped quotes are handled correctly regardless of format
        let workbook2 = Workbook(fromCSV: csv, hasHeader: true)
        let sheet2 = try #require(workbook2.getSheets().first)
        
        #expect(sheet2.getCell("A2") == .string("Apple"))
        #expect(sheet2.getCell("B2") == .string("Red, delicious"))
        #expect(sheet2.getCell("A3") == .string("Quote"))
        // Verify that quotes are preserved correctly in the round-trip
        let quoteValue = sheet2.getCell("B3")
        #expect(quoteValue == .string("\"Hello\""), "Quoted strings should be preserved correctly in CSV round-trip")
    }
    
    @Test func testCSVWithEmptyFields() throws {
        // Test CSV with empty fields
        let csvData = "Name,Age,City\nAlice,30,\nBob,,Paris\n,25,London"
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        let sheet = try #require(workbook.getSheets().first)
        
        #expect(sheet.getCell("A2") == .string("Alice"))
        #expect(sheet.getCell("B2") == .integer(30))
        #expect(sheet.getCell("C2") == .string(""))
        
        #expect(sheet.getCell("A3") == .string("Bob"))
        #expect(sheet.getCell("B3") == .string(""))
        #expect(sheet.getCell("C3") == .string("Paris"))
        
        #expect(sheet.getCell("A4") == .string(""))
        #expect(sheet.getCell("B4") == .integer(25))
        #expect(sheet.getCell("C4") == .string("London"))
    }
}
