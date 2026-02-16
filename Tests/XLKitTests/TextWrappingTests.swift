//
//  TextWrappingTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit
@testable import XLKitXLSX

@MainActor
final class TextWrappingTests: XLKitTestBase {
    
    func testTextWrapping() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Text Wrapping Test")
        
        // Test text wrapping enabled
        var wrappedFormat = CellFormat()
        wrappedFormat.textWrapping = true
        wrappedFormat.verticalAlignment = .top
        
        sheet.setCell("A1", string: "This is a long text that should wrap to multiple lines when text wrapping is enabled.", format: wrappedFormat)
        
        // Test text wrapping disabled
        var noWrapFormat = CellFormat()
        noWrapFormat.textWrapping = false
        noWrapFormat.verticalAlignment = .top
        
        sheet.setCell("A2", string: "This text should not wrap and may be truncated.", format: noWrapFormat)
        
        // Test text wrapping with other formatting
        var wrappedBoldFormat = CellFormat()
        wrappedBoldFormat.textWrapping = true
        wrappedBoldFormat.verticalAlignment = .top
        wrappedBoldFormat.fontWeight = .bold
        wrappedBoldFormat.backgroundColor = "#E6E6E6"
        
        sheet.setCell("A3", string: "Bold wrapped text with background color.", format: wrappedBoldFormat)
        
        // Set column width to demonstrate wrapping
        sheet.setColumnWidth(1, width: 200)
        
        // Save and verify
        let tempURL = makeTempWorkbookURL(prefix: "text_wrapping_test")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with text wrapping: \(error)")
        }
    }
    
    func testTextWrappingInFormatToKey() {
        var format1 = CellFormat()
        format1.textWrapping = true
        
        var format2 = CellFormat()
        format2.textWrapping = false
        
        let format3 = CellFormat() // nil textWrapping
        
        let key1 = XLSXEngine.formatToKey(format1)
        let key2 = XLSXEngine.formatToKey(format2)
        let key3 = XLSXEngine.formatToKey(format3)
        
        XCTAssertTrue(key1.contains("textWrapping:true"), "Text wrapping true should be in format key")
        XCTAssertTrue(key2.contains("textWrapping:false"), "Text wrapping false should be in format key")
        XCTAssertTrue(key3.contains("textWrapping:false"), "Nil text wrapping should default to false in format key")
    }
}
