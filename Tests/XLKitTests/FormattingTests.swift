//
//  FormattingTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct FormattingTests {
    
    @Test func testCellFormatting() {
        let headerFormat = CellFormat.header()
        let currencyFormat = CellFormat.currency()
        let percentageFormat = CellFormat.percentage()
        let dateFormat = CellFormat.date()
        let borderedFormat = CellFormat.bordered()
        
        // Factory helpers return non-optional CellFormat values; verify they are usable.
        #expect(headerFormat.fontWeight == .bold)
        #expect(currencyFormat.numberFormat != nil)
        #expect(percentageFormat.numberFormat != nil)
        #expect(dateFormat.numberFormat != nil)
        #expect(borderedFormat.borderTop != nil)
    }
    
    @Test func testFontColorFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Font Color Test")
        
        // Test different font colors
        let redTextFormat = CellFormat.text(color: "#FF0000")
        let blueTextFormat = CellFormat.text(color: "#0000FF")
        let greenTextFormat = CellFormat.text(color: "#00FF00")
        let currencyRedFormat = CellFormat.currency(color: "#FF0000")
        
        // Set cells with different colors
        sheet.setCell("A1", string: "Red Text", format: redTextFormat)
        sheet.setCell("A2", string: "Blue Text", format: blueTextFormat)
        sheet.setCell("A3", string: "Green Text", format: greenTextFormat)
        sheet.setCell("A4", number: 1234.56, format: currencyRedFormat)
        
        // Verify formats are applied
        let redCell = sheet.getCellWithFormat("A1")
        let blueCell = sheet.getCellWithFormat("A2")
        let greenCell = sheet.getCellWithFormat("A3")
        let currencyCell = sheet.getCellWithFormat("A4")
        
        #expect(redCell != nil)
        #expect(blueCell != nil)
        #expect(greenCell != nil)
        #expect(currencyCell != nil)
        
        #expect(redCell?.format?.fontColor == "#FF0000")
        #expect(blueCell?.format?.fontColor == "#0000FF")
        #expect(greenCell?.format?.fontColor == "#00FF00")
        #expect(currencyCell?.format?.fontColor == "#FF0000")
    }
    
    @Test func testHorizontalAlignment() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Horizontal Alignment Test")
        
        // Test all horizontal alignment options
        var leftAlignFormat = CellFormat()
        leftAlignFormat.horizontalAlignment = .left
        
        var centerAlignFormat = CellFormat()
        centerAlignFormat.horizontalAlignment = .center
        
        var rightAlignFormat = CellFormat()
        rightAlignFormat.horizontalAlignment = .right
        
        var justifyAlignFormat = CellFormat()
        justifyAlignFormat.horizontalAlignment = .justify
        
        var distributedAlignFormat = CellFormat()
        distributedAlignFormat.horizontalAlignment = .distributed
        
        // Set cells with different horizontal alignments
        sheet.setCell("A1", string: "Left Aligned", format: leftAlignFormat)
        sheet.setCell("A2", string: "Center Aligned", format: centerAlignFormat)
        sheet.setCell("A3", string: "Right Aligned", format: rightAlignFormat)
        sheet.setCell("A4", string: "Justify Aligned", format: justifyAlignFormat)
        sheet.setCell("A5", string: "Distributed Aligned", format: distributedAlignFormat)
        
        // Verify alignments are applied
        let leftCell = sheet.getCellWithFormat("A1")
        let centerCell = sheet.getCellWithFormat("A2")
        let rightCell = sheet.getCellWithFormat("A3")
        let justifyCell = sheet.getCellWithFormat("A4")
        let distributedCell = sheet.getCellWithFormat("A5")
        
        #expect(leftCell != nil)
        #expect(centerCell != nil)
        #expect(rightCell != nil)
        #expect(justifyCell != nil)
        #expect(distributedCell != nil)
        
        #expect(leftCell?.format?.horizontalAlignment == .left)
        #expect(centerCell?.format?.horizontalAlignment == .center)
        #expect(rightCell?.format?.horizontalAlignment == .right)
        #expect(justifyCell?.format?.horizontalAlignment == .justify)
        #expect(distributedCell?.format?.horizontalAlignment == .distributed)
    }
    
    @Test func testVerticalAlignment() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Vertical Alignment Test")
        
        // Test all vertical alignment options
        var topAlignFormat = CellFormat()
        topAlignFormat.verticalAlignment = .top
        
        var centerAlignFormat = CellFormat()
        centerAlignFormat.verticalAlignment = .center
        
        var bottomAlignFormat = CellFormat()
        bottomAlignFormat.verticalAlignment = .bottom
        
        var justifyAlignFormat = CellFormat()
        justifyAlignFormat.verticalAlignment = .justify
        
        var distributedAlignFormat = CellFormat()
        distributedAlignFormat.verticalAlignment = .distributed
        
        // Set cells with different vertical alignments
        sheet.setCell("A1", string: "Top Aligned", format: topAlignFormat)
        sheet.setCell("A2", string: "Center Aligned", format: centerAlignFormat)
        sheet.setCell("A3", string: "Bottom Aligned", format: bottomAlignFormat)
        sheet.setCell("A4", string: "Justify Aligned", format: justifyAlignFormat)
        sheet.setCell("A5", string: "Distributed Aligned", format: distributedAlignFormat)
        
        // Verify alignments are applied
        let topCell = sheet.getCellWithFormat("A1")
        let centerCell = sheet.getCellWithFormat("A2")
        let bottomCell = sheet.getCellWithFormat("A3")
        let justifyCell = sheet.getCellWithFormat("A4")
        let distributedCell = sheet.getCellWithFormat("A5")
        
        #expect(topCell != nil)
        #expect(centerCell != nil)
        #expect(bottomCell != nil)
        #expect(justifyCell != nil)
        #expect(distributedCell != nil)
        
        #expect(topCell?.format?.verticalAlignment == .top)
        #expect(centerCell?.format?.verticalAlignment == .center)
        #expect(bottomCell?.format?.verticalAlignment == .bottom)
        #expect(justifyCell?.format?.verticalAlignment == .justify)
        #expect(distributedCell?.format?.verticalAlignment == .distributed)
    }
    
    @Test func testCombinedAlignment() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Combined Alignment Test")
        
        // Test combinations of horizontal and vertical alignments
        var topLeftFormat = CellFormat()
        topLeftFormat.horizontalAlignment = .left
        topLeftFormat.verticalAlignment = .top
        
        var centerCenterFormat = CellFormat()
        centerCenterFormat.horizontalAlignment = .center
        centerCenterFormat.verticalAlignment = .center
        
        var bottomRightFormat = CellFormat()
        bottomRightFormat.horizontalAlignment = .right
        bottomRightFormat.verticalAlignment = .bottom
        
        var justifyDistributedFormat = CellFormat()
        justifyDistributedFormat.horizontalAlignment = .justify
        justifyDistributedFormat.verticalAlignment = .distributed
        
        // Set cells with combined alignments
        sheet.setCell("A1", string: "Top-Left", format: topLeftFormat)
        sheet.setCell("A2", string: "Center-Center", format: centerCenterFormat)
        sheet.setCell("A3", string: "Bottom-Right", format: bottomRightFormat)
        sheet.setCell("A4", string: "Justify-Distributed", format: justifyDistributedFormat)
        
        // Verify combined alignments are applied
        let topLeftCell = sheet.getCellWithFormat("A1")
        let centerCenterCell = sheet.getCellWithFormat("A2")
        let bottomRightCell = sheet.getCellWithFormat("A3")
        let justifyDistributedCell = sheet.getCellWithFormat("A4")
        
        #expect(topLeftCell != nil)
        #expect(centerCenterCell != nil)
        #expect(bottomRightCell != nil)
        #expect(justifyDistributedCell != nil)
        
        // Test horizontal alignments
        #expect(topLeftCell?.format?.horizontalAlignment == .left)
        #expect(centerCenterCell?.format?.horizontalAlignment == .center)
        #expect(bottomRightCell?.format?.horizontalAlignment == .right)
        #expect(justifyDistributedCell?.format?.horizontalAlignment == .justify)
        
        // Test vertical alignments
        #expect(topLeftCell?.format?.verticalAlignment == .top)
        #expect(centerCenterCell?.format?.verticalAlignment == .center)
        #expect(bottomRightCell?.format?.verticalAlignment == .bottom)
        #expect(justifyDistributedCell?.format?.verticalAlignment == .distributed)
    }
    
    @Test func testAlignmentWithOtherFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Alignment with Formatting Test")
        
        // Test alignment combined with other formatting options
        var formattedAlignFormat = CellFormat()
        formattedAlignFormat.horizontalAlignment = .center
        formattedAlignFormat.verticalAlignment = .center
        formattedAlignFormat.fontWeight = .bold
        formattedAlignFormat.fontSize = 14.0
        formattedAlignFormat.fontColor = "#FF0000"
        formattedAlignFormat.backgroundColor = "#FFFF00"
        
        // Set cell with combined formatting
        sheet.setCell("A1", string: "Formatted & Aligned", format: formattedAlignFormat)
        
        // Verify all formatting is applied
        let formattedCell = sheet.getCellWithFormat("A1")
        #expect(formattedCell != nil)
        
        #expect(formattedCell?.format?.horizontalAlignment == .center)
        #expect(formattedCell?.format?.verticalAlignment == .center)
        #expect(formattedCell?.format?.fontWeight == .bold)
        #expect(formattedCell?.format?.fontSize == 14.0)
        #expect(formattedCell?.format?.fontColor == "#FF0000")
        #expect(formattedCell?.format?.backgroundColor == "#FFFF00")
        // Verify text wrapping remains at the expected default (nil, which defaults to false in format key)
        #expect(formattedCell?.format?.textWrapping == nil)
    }
    
    @Test func testAlignmentEnumValues() {
        // Test that all alignment enum values are correct
        #expect(HorizontalAlignment.left.rawValue == "left")
        #expect(HorizontalAlignment.center.rawValue == "center")
        #expect(HorizontalAlignment.right.rawValue == "right")
        #expect(HorizontalAlignment.justify.rawValue == "justify")
        #expect(HorizontalAlignment.distributed.rawValue == "distributed")
        
        #expect(VerticalAlignment.top.rawValue == "top")
        #expect(VerticalAlignment.center.rawValue == "center")
        #expect(VerticalAlignment.bottom.rawValue == "bottom")
        #expect(VerticalAlignment.justify.rawValue == "justify")
        #expect(VerticalAlignment.distributed.rawValue == "distributed")
    }
    
    @Test func testSetCellWithFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let headerCell = Cell.string("Header", format: CellFormat.header())
        let currencyCell = Cell.number(1234.56, format: CellFormat.currency())
        
        sheet.setCell("A1", cell: headerCell)
        sheet.setCell("B1", cell: currencyCell)
        
        let retrievedHeader = sheet.getCellWithFormat("A1")
        let retrievedCurrency = sheet.getCellWithFormat("B1")
        
        #expect(retrievedHeader != nil)
        #expect(retrievedCurrency != nil)
        #expect(retrievedHeader?.value == .string("Header"))
        #expect(retrievedCurrency?.value == .number(1234.56))
    }
}
