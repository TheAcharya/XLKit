//
//  BorderTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct BorderTests {
    
    @Test func testBordersActuallyWork() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border Test")
        
        // Test different border styles
        let thinBorderFormat = XLKitTestSupport.makeThinBorderFormat()
        let mediumBorderFormat = XLKitTestSupport.makeMediumRedBorderFormat()
        let thickBorderFormat = XLKitTestSupport.makeThickBlueBorderFormat()
        
        // Set cells with different border styles
        sheet.setCell("A1", string: "Thin Borders", format: thinBorderFormat)
        sheet.setCell("A2", string: "Medium Red Borders", format: mediumBorderFormat)
        sheet.setCell("A3", string: "Thick Blue Borders", format: thickBorderFormat)
        
        // Verify border formats are stored correctly
        let thinCell = sheet.getCellWithFormat("A1")
        let mediumCell = sheet.getCellWithFormat("A2")
        let thickCell = sheet.getCellWithFormat("A3")
        
        #expect(thinCell != nil)
        #expect(mediumCell != nil)
        #expect(thickCell != nil)
        
        #expect(thinCell?.format?.borderTop == .thin)
        #expect(thinCell?.format?.borderBottom == .thin)
        #expect(thinCell?.format?.borderLeft == .thin)
        #expect(thinCell?.format?.borderRight == .thin)
        
        #expect(mediumCell?.format?.borderTop == .medium)
        #expect(mediumCell?.format?.borderBottom == .medium)
        #expect(mediumCell?.format?.borderColor == "#FF0000")
        
        #expect(thickCell?.format?.borderTop == .thick)
        #expect(thickCell?.format?.borderBottom == .thick)
        #expect(thickCell?.format?.borderColor == "#0000FF")
        
        // Test Excel file generation with borders
        let tempURL = XLKitTestSupport.makeTempWorkbookURL(prefix: "border_test")
        
        try workbook.save(to: tempURL)
        #expect(FileManager.default.fileExists(atPath: tempURL.path))
        try FileManager.default.removeItem(at: tempURL)
    }
    
    @Test func testDifferentBorderStyles() {
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
            #expect(cell != nil)
            #expect(cell?.format?.borderTop == style)
            #expect(cell?.format?.borderBottom == style)
            #expect(cell?.format?.borderLeft == style)
            #expect(cell?.format?.borderRight == style)
            #expect(cell?.format?.borderColor == borderColors[index])
        }
    }
    
    @Test func testBorderWithOtherFormatting() {
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
        #expect(cell != nil)
        
        #expect(cell?.format?.borderTop == .thick)
        #expect(cell?.format?.borderBottom == .thick)
        #expect(cell?.format?.borderColor == "#FF0000")
        #expect(cell?.format?.fontWeight == .bold)
        #expect(cell?.format?.fontSize == 14.0)
        #expect(cell?.format?.fontColor == "#0000FF")
        #expect(cell?.format?.backgroundColor == "#FFFF00")
        #expect(cell?.format?.horizontalAlignment == .center)
        #expect(cell?.format?.verticalAlignment == .center)
    }
}
