//
//  NumberFormatTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit
@testable import XLKitXLSX

@Suite
@MainActor
struct NumberFormatTests {
    
    @Test func testNumberFormatCurrency() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Test currency formatting
        let currencyFormat = CellFormat.currency()
        sheet.setCell("A1", number: 1234.56, format: currencyFormat)
        
        // Verify the format is stored correctly
        let storedFormat = sheet.getCellFormat("A1")
        #expect(storedFormat != nil)
        #expect(storedFormat?.numberFormat == .currencyWithDecimals)
        
        // Test custom currency format
        var customCurrencyFormat = CellFormat.currency()
        customCurrencyFormat.numberFormat = .custom
        customCurrencyFormat.customNumberFormat = "$#,##0.00"
        sheet.setCell("A2", number: 5678.90, format: customCurrencyFormat)
        
        let storedCustomFormat = sheet.getCellFormat("A2")
        #expect(storedCustomFormat != nil)
        #expect(storedCustomFormat?.numberFormat == .custom)
        #expect(storedCustomFormat?.customNumberFormat == "$#,##0.00")
    }
    
    @Test func testNumberFormatPercentage() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Test percentage formatting
        let percentageFormat = CellFormat.percentage()
        sheet.setCell("A1", number: 0.1234, format: percentageFormat)
        
        // Verify the format is stored correctly
        let storedFormat = sheet.getCellFormat("A1")
        #expect(storedFormat != nil)
        #expect(storedFormat?.numberFormat == .percentageWithDecimals)
    }
    
    @Test func testNumberFormatCustom() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Test custom number format
        var customFormat = CellFormat()
        customFormat.numberFormat = .custom
        customFormat.customNumberFormat = "#,##0.00"
        
        sheet.setCell("A1", number: 1234.56, format: customFormat)
        
        // Verify the format is stored correctly
        let storedFormat = sheet.getCellFormat("A1")
        #expect(storedFormat != nil)
        #expect(storedFormat?.numberFormat == .custom)
        #expect(storedFormat?.customNumberFormat == "#,##0.00")
    }
    
    @Test func testNumberFormatInFormatToKey() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Create formats with different number formats
        let currencyFormat = CellFormat.currency()
        let percentageFormat = CellFormat.percentage()
        var customFormat = CellFormat()
        customFormat.numberFormat = .custom
        customFormat.customNumberFormat = "#,##0.00"
        
        // Set cells with different formats
        sheet.setCell("A1", number: 1234.56, format: currencyFormat)
        sheet.setCell("A2", number: 0.1234, format: percentageFormat)
        sheet.setCell("A3", number: 5678.90, format: customFormat)
        
        // Verify that formatToKey generates different keys for different number formats
        let key1 = XLSXEngine.formatToKey(currencyFormat)
        let key2 = XLSXEngine.formatToKey(percentageFormat)
        let key3 = XLSXEngine.formatToKey(customFormat)
        
        #expect(key1 != key2, "Currency and percentage formats should have different keys")
        #expect(key1 != key3, "Currency and custom formats should have different keys")
        #expect(key2 != key3, "Percentage and custom formats should have different keys")
        
        // Verify that number format information is included in the key
        #expect(key1.contains("numberFormat:$#,##0.00"), "Currency format key should include number format")
        #expect(key2.contains("numberFormat:0.00%"), "Percentage format key should include number format")
        #expect(key3.contains("numberFormat:#,##0.00"), "Custom format key should include number format")
    }
    
    @Test func testNumberFormatExcelGeneration() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Add cells with different number formats
        sheet.setCell("A1", number: 1234.56, format: .currency())
        sheet.setCell("A2", number: 0.1234, format: .percentage())
        
        var customFormat = CellFormat()
        customFormat.numberFormat = .custom
        customFormat.customNumberFormat = "$#,##0.00"
        sheet.setCell("A3", number: 5678.90, format: customFormat)
        
        // Save the workbook to test XLSX generation
        let tempURL = XLKitTestSupport.makeTempWorkbookURL(prefix: "number_format_test")
        defer { XLKitTestSupport.cleanupTempFile(at: tempURL) }
        
        try workbook.save(to: tempURL)
        #expect(FileManager.default.fileExists(atPath: tempURL.path))
    }
}
