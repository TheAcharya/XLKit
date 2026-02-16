//
//  BorderTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class BorderTests: XLKitTestBase {
    
    func testBordersActuallyWork() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border Test")
        
        // Test different border styles
        let thinBorderFormat = Self.makeThinBorderFormat()
        let mediumBorderFormat = Self.makeMediumRedBorderFormat()
        let thickBorderFormat = Self.makeThickBlueBorderFormat()
        
        // Set cells with different border styles
        sheet.setCell("A1", string: "Thin Borders", format: thinBorderFormat)
        sheet.setCell("A2", string: "Medium Red Borders", format: mediumBorderFormat)
        sheet.setCell("A3", string: "Thick Blue Borders", format: thickBorderFormat)
        
        // Verify border formats are stored correctly
        let thinCell = sheet.getCellWithFormat("A1")
        let mediumCell = sheet.getCellWithFormat("A2")
        let thickCell = sheet.getCellWithFormat("A3")
        
        XCTAssertNotNil(thinCell)
        XCTAssertNotNil(mediumCell)
        XCTAssertNotNil(thickCell)
        
        XCTAssertEqual(thinCell?.format?.borderTop, .thin)
        XCTAssertEqual(thinCell?.format?.borderBottom, .thin)
        XCTAssertEqual(thinCell?.format?.borderLeft, .thin)
        XCTAssertEqual(thinCell?.format?.borderRight, .thin)
        
        XCTAssertEqual(mediumCell?.format?.borderTop, .medium)
        XCTAssertEqual(mediumCell?.format?.borderBottom, .medium)
        XCTAssertEqual(mediumCell?.format?.borderColor, "#FF0000")
        
        XCTAssertEqual(thickCell?.format?.borderTop, .thick)
        XCTAssertEqual(thickCell?.format?.borderBottom, .thick)
        XCTAssertEqual(thickCell?.format?.borderColor, "#0000FF")
        
        // Test Excel file generation with borders
        let tempURL = makeTempWorkbookURL(prefix: "border_test")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with borders: \(error)")
        }
    }
    
    func testDifferentBorderStyles() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border Styles Test")
        
        // Test all border styles
        let borderStyles: [BorderStyle] = [.thin, .medium, .thick]
        let borderColors = ["#FF0000", "#00FF00", "#0000FF"]
        
        for (index, style) in borderStyles.enumerated() {
            var borderFormat = CellFormat.bordered()
            borderFormat.borderTop = style
            borderFormat.borderBottom = style
            borderFormat.borderLeft = style
            borderFormat.borderRight = style
            borderFormat.borderColor = borderColors[index]
            
            let cellAddress = "A\(index + 1)"
            sheet.setCell(cellAddress, string: "Style: \(style.rawValue)", format: borderFormat)
            
            // Verify border style is stored correctly
            let cell = sheet.getCellWithFormat(cellAddress)
            XCTAssertNotNil(cell)
            XCTAssertEqual(cell?.format?.borderTop, style)
            XCTAssertEqual(cell?.format?.borderBottom, style)
            XCTAssertEqual(cell?.format?.borderLeft, style)
            XCTAssertEqual(cell?.format?.borderRight, style)
            XCTAssertEqual(cell?.format?.borderColor, borderColors[index])
        }
    }
    
    func testBorderWithOtherFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border with Formatting Test")
        
        // Test borders combined with other formatting
        var formattedBorderFormat = CellFormat.bordered()
        formattedBorderFormat.borderTop = .thick
        formattedBorderFormat.borderBottom = .thick
        formattedBorderFormat.borderColor = "#FF0000"
        formattedBorderFormat.fontWeight = .bold
        formattedBorderFormat.fontSize = 14.0
        formattedBorderFormat.fontColor = "#0000FF"
        formattedBorderFormat.backgroundColor = "#FFFF00"
        formattedBorderFormat.horizontalAlignment = .center
        formattedBorderFormat.verticalAlignment = .center
        
        sheet.setCell("A1", string: "Formatted with Borders", format: formattedBorderFormat)
        
        // Verify all formatting is applied
        let cell = sheet.getCellWithFormat("A1")
        XCTAssertNotNil(cell)
        
        XCTAssertEqual(cell?.format?.borderTop, .thick)
        XCTAssertEqual(cell?.format?.borderBottom, .thick)
        XCTAssertEqual(cell?.format?.borderColor, "#FF0000")
        XCTAssertEqual(cell?.format?.fontWeight, .bold)
        XCTAssertEqual(cell?.format?.fontSize, 14.0)
        XCTAssertEqual(cell?.format?.fontColor, "#0000FF")
        XCTAssertEqual(cell?.format?.backgroundColor, "#FFFF00")
        XCTAssertEqual(cell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(cell?.format?.verticalAlignment, .center)
    }
}
