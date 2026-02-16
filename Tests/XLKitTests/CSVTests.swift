//
//  CSVTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class CSVTests: XLKitTestBase {
    
    func testWorkbookFromCSV() {
        let csvData = """
        Name,Age,Salary
        Alice,30,50000.5
        Bob,25,45000.75
        """
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        let sheets = workbook.getSheets()
        XCTAssertTrue(!sheets.isEmpty, "Workbook created from CSV should contain at least one sheet")
        guard let sheet = sheets.first else {
            XCTFail("Expected workbook created from CSV to contain at least one sheet")
            return
        }
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Alice"))
        XCTAssertEqual(sheet.getCell("B2"), .integer(30))
        XCTAssertEqual(sheet.getCell("C2"), .number(50000.5))
        XCTAssertEqual(sheet.getCell("A3"), .string("Bob"))
        XCTAssertEqual(sheet.getCell("B3"), .integer(25))
        XCTAssertEqual(sheet.getCell("C3"), .number(45000.75))
    }
    
    func testWorkbookFromTSV() {
        let tsvData = """
        Product\tPrice\tIn Stock
        Apple\t1.99\ttrue
        Banana\t0.99\tfalse
        """
        
        let workbook = Workbook(fromTSV: tsvData, hasHeader: true)
        guard let sheet = workbook.getSheets().first else {
            XCTFail("Expected workbook created from TSV to contain at least one sheet")
            return
        }
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Apple"))
        XCTAssertEqual(sheet.getCell("B2"), .number(1.99))
        XCTAssertEqual(sheet.getCell("C2"), .boolean(true))
        XCTAssertEqual(sheet.getCell("A3"), .string("Banana"))
        XCTAssertEqual(sheet.getCell("B3"), .number(0.99))
        XCTAssertEqual(sheet.getCell("C3"), .boolean(false))
    }
    
    func testSheetExportToCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Age", "Salary"])
        sheet.setRow(2, strings: ["Alice", "30", "50000"])
        sheet.setRow(3, strings: ["Bob", "25", "45000"])
        
        let csv = sheet.exportToCSV()
        let expectedCSV = "Name,Age,Salary\nAlice,30,50000\nBob,25,45000"
        
        XCTAssertEqual(csv, expectedCSV)
    }
    
    func testSheetExportToTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Product", "Price", "In Stock"])
        sheet.setRow(2, strings: ["Apple", "1.99", "true"])
        sheet.setRow(3, strings: ["Banana", "0.99", "false"])
        
        let tsv = sheet.exportToTSV()
        let expectedTSV = "Product\tPrice\tIn Stock\nApple\t1.99\ttrue\nBanana\t0.99\tfalse"
        
        XCTAssertEqual(tsv, expectedTSV)
    }
    
    func testWorkbookImportCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let csvData = """
        Name,Age,Salary
        Alice,30,50000.5
        Bob,25,45000.75
        """
        
        workbook.importCSV(csvData, into: sheet, hasHeader: true)
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Alice"))
        XCTAssertEqual(sheet.getCell("B2"), .integer(30))
        XCTAssertEqual(sheet.getCell("C2"), .number(50000.5))
        XCTAssertEqual(sheet.getCell("A3"), .string("Bob"))
        XCTAssertEqual(sheet.getCell("B3"), .integer(25))
        XCTAssertEqual(sheet.getCell("C3"), .number(45000.75))
    }
    
    func testWorkbookImportTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let tsvData = """
        Product\tPrice\tIn Stock
        Apple\t1.99\ttrue
        Banana\t0.99\tfalse
        """
        
        workbook.importTSV(tsvData, into: sheet, hasHeader: true)
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Apple"))
        XCTAssertEqual(sheet.getCell("B2"), .number(1.99))
        XCTAssertEqual(sheet.getCell("C2"), .boolean(true))
        XCTAssertEqual(sheet.getCell("A3"), .string("Banana"))
        XCTAssertEqual(sheet.getCell("B3"), .number(0.99))
        XCTAssertEqual(sheet.getCell("C3"), .boolean(false))
    }
    
    func testWorkbookExportSheetToCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Age", "Salary"])
        sheet.setRow(2, strings: ["Alice", "30", "50000"])
        sheet.setRow(3, strings: ["Bob", "25", "45000"])
        
        let csv = workbook.exportSheetToCSV(sheet)
        let expectedCSV = "Name,Age,Salary\nAlice,30,50000\nBob,25,45000"
        
        XCTAssertEqual(csv, expectedCSV)
    }
    
    func testWorkbookExportSheetToTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Product", "Price", "In Stock"])
        sheet.setRow(2, strings: ["Apple", "1.99", "true"])
        sheet.setRow(3, strings: ["Banana", "0.99", "false"])
        
        let tsv = workbook.exportSheetToTSV(sheet)
        let expectedTSV = "Product\tPrice\tIn Stock\nApple\t1.99\ttrue\nBanana\t0.99\tfalse"
        
        XCTAssertEqual(tsv, expectedTSV)
    }
    
    func testCSVWithQuotedFields() {
        // Test CSV with quoted fields containing commas (swift-textfile-tools handles this)
        let csvData = "Name,Description,Price\nApple,\"Red, delicious apple\",1.99\nBanana,\"Yellow, curved fruit\",0.99"
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        guard let sheet = workbook.getSheets().first else {
            XCTFail("Expected workbook created from CSV to contain at least one sheet")
            return
        }
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Apple"))
        XCTAssertEqual(sheet.getCell("B2"), .string("Red, delicious apple"))
        XCTAssertEqual(sheet.getCell("C2"), .number(1.99))
        XCTAssertEqual(sheet.getCell("A3"), .string("Banana"))
        XCTAssertEqual(sheet.getCell("B3"), .string("Yellow, curved fruit"))
        XCTAssertEqual(sheet.getCell("C3"), .number(0.99))
    }
    
    func testCSVWithEscapedQuotes() {
        // Test CSV with escaped quotes (double quotes) - swift-textfile-tools handles this
        let csvData = "Name,Quote\nAlice,\"She said \"\"Hello\"\"\"\nBob,\"He said \"\"Goodbye\"\"\""
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        guard let sheet = workbook.getSheets().first else {
            XCTFail("Expected workbook created from CSV to contain at least one sheet")
            return
        }
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Alice"))
        XCTAssertEqual(sheet.getCell("B2"), .string("She said \"Hello\""))
        XCTAssertEqual(sheet.getCell("A3"), .string("Bob"))
        XCTAssertEqual(sheet.getCell("B3"), .string("He said \"Goodbye\""))
    }
    
    func testCSVExportWithSpecialCharacters() {
        // Test that CSV export properly quotes fields with special characters
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Description"])
        sheet.setRow(2, strings: ["Apple", "Red, delicious"])
        sheet.setRow(3, strings: ["Quote", "\"Hello\""])
        
        let csv = sheet.exportToCSV()
        
        // Verify the CSV contains quoted fields for special characters
        XCTAssertTrue(csv.contains("\"Red, delicious\""), "CSV should quote fields containing commas")
        
        // Verify round-trip: import the exported CSV (this validates CSV format correctness)
        // The round-trip test ensures escaped quotes are handled correctly regardless of format
        let workbook2 = Workbook(fromCSV: csv, hasHeader: true)
        guard let sheet2 = workbook2.getSheets().first else {
            XCTFail("Expected workbook created from CSV to contain at least one sheet")
            return
        }
        
        XCTAssertEqual(sheet2.getCell("A2"), .string("Apple"))
        XCTAssertEqual(sheet2.getCell("B2"), .string("Red, delicious"))
        XCTAssertEqual(sheet2.getCell("A3"), .string("Quote"))
        // Verify that quotes are preserved correctly in the round-trip
        let quoteValue = sheet2.getCell("B3")
        XCTAssertEqual(quoteValue, .string("\"Hello\""), "Quoted strings should be preserved correctly in CSV round-trip")
    }
    
    func testCSVWithEmptyFields() {
        // Test CSV with empty fields
        let csvData = "Name,Age,City\nAlice,30,\nBob,,Paris\n,25,London"
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        guard let sheet = workbook.getSheets().first else {
            XCTFail("Expected workbook created from CSV to contain at least one sheet")
            return
        }
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Alice"))
        XCTAssertEqual(sheet.getCell("B2"), .integer(30))
        XCTAssertEqual(sheet.getCell("C2"), .string(""))
        
        XCTAssertEqual(sheet.getCell("A3"), .string("Bob"))
        XCTAssertEqual(sheet.getCell("B3"), .string(""))
        XCTAssertEqual(sheet.getCell("C3"), .string("Paris"))
        
        XCTAssertEqual(sheet.getCell("A4"), .string(""))
        XCTAssertEqual(sheet.getCell("B4"), .integer(25))
        XCTAssertEqual(sheet.getCell("C4"), .string("London"))
    }
}
