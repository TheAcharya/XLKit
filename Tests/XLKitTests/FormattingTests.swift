//
//  FormattingTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class FormattingTests: XLKitTestBase {
    
    func testCellFormatting() {
        let headerFormat = CellFormat.header()
        let currencyFormat = CellFormat.currency()
        let percentageFormat = CellFormat.percentage()
        let dateFormat = CellFormat.date()
        let borderedFormat = CellFormat.bordered()
        
        XCTAssertNotNil(headerFormat)
        XCTAssertNotNil(currencyFormat)
        XCTAssertNotNil(percentageFormat)
        XCTAssertNotNil(dateFormat)
        XCTAssertNotNil(borderedFormat)
    }
    
    func testFontColorFormatting() {
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
        
        XCTAssertNotNil(redCell)
        XCTAssertNotNil(blueCell)
        XCTAssertNotNil(greenCell)
        XCTAssertNotNil(currencyCell)
        
        XCTAssertEqual(redCell?.format?.fontColor, "#FF0000")
        XCTAssertEqual(blueCell?.format?.fontColor, "#0000FF")
        XCTAssertEqual(greenCell?.format?.fontColor, "#00FF00")
        XCTAssertEqual(currencyCell?.format?.fontColor, "#FF0000")
    }
    
    func testHorizontalAlignment() {
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
        
        XCTAssertNotNil(leftCell)
        XCTAssertNotNil(centerCell)
        XCTAssertNotNil(rightCell)
        XCTAssertNotNil(justifyCell)
        XCTAssertNotNil(distributedCell)
        
        XCTAssertEqual(leftCell?.format?.horizontalAlignment, .left)
        XCTAssertEqual(centerCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(rightCell?.format?.horizontalAlignment, .right)
        XCTAssertEqual(justifyCell?.format?.horizontalAlignment, .justify)
        XCTAssertEqual(distributedCell?.format?.horizontalAlignment, .distributed)
    }
    
    func testVerticalAlignment() {
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
        
        XCTAssertNotNil(topCell)
        XCTAssertNotNil(centerCell)
        XCTAssertNotNil(bottomCell)
        XCTAssertNotNil(justifyCell)
        XCTAssertNotNil(distributedCell)
        
        XCTAssertEqual(topCell?.format?.verticalAlignment, .top)
        XCTAssertEqual(centerCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(bottomCell?.format?.verticalAlignment, .bottom)
        XCTAssertEqual(justifyCell?.format?.verticalAlignment, .justify)
        XCTAssertEqual(distributedCell?.format?.verticalAlignment, .distributed)
    }
    
    func testCombinedAlignment() {
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
        
        XCTAssertNotNil(topLeftCell)
        XCTAssertNotNil(centerCenterCell)
        XCTAssertNotNil(bottomRightCell)
        XCTAssertNotNil(justifyDistributedCell)
        
        // Test horizontal alignments
        XCTAssertEqual(topLeftCell?.format?.horizontalAlignment, .left)
        XCTAssertEqual(centerCenterCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(bottomRightCell?.format?.horizontalAlignment, .right)
        XCTAssertEqual(justifyDistributedCell?.format?.horizontalAlignment, .justify)
        
        // Test vertical alignments
        XCTAssertEqual(topLeftCell?.format?.verticalAlignment, .top)
        XCTAssertEqual(centerCenterCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(bottomRightCell?.format?.verticalAlignment, .bottom)
        XCTAssertEqual(justifyDistributedCell?.format?.verticalAlignment, .distributed)
    }
    
    func testAlignmentWithOtherFormatting() {
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
        XCTAssertNotNil(formattedCell)
        
        XCTAssertEqual(formattedCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(formattedCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(formattedCell?.format?.fontWeight, .bold)
        XCTAssertEqual(formattedCell?.format?.fontSize, 14.0)
        XCTAssertEqual(formattedCell?.format?.fontColor, "#FF0000")
        XCTAssertEqual(formattedCell?.format?.backgroundColor, "#FFFF00")
        // Verify text wrapping remains at the expected default (nil, which defaults to false in format key)
        XCTAssertNil(formattedCell?.format?.textWrapping)
    }
    
    func testAlignmentEnumValues() {
        // Test that all alignment enum values are correct
        XCTAssertEqual(HorizontalAlignment.left.rawValue, "left")
        XCTAssertEqual(HorizontalAlignment.center.rawValue, "center")
        XCTAssertEqual(HorizontalAlignment.right.rawValue, "right")
        XCTAssertEqual(HorizontalAlignment.justify.rawValue, "justify")
        XCTAssertEqual(HorizontalAlignment.distributed.rawValue, "distributed")
        
        XCTAssertEqual(VerticalAlignment.top.rawValue, "top")
        XCTAssertEqual(VerticalAlignment.center.rawValue, "center")
        XCTAssertEqual(VerticalAlignment.bottom.rawValue, "bottom")
        XCTAssertEqual(VerticalAlignment.justify.rawValue, "justify")
        XCTAssertEqual(VerticalAlignment.distributed.rawValue, "distributed")
    }
    
    func testSetCellWithFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let headerCell = Cell.string("Header", format: CellFormat.header())
        let currencyCell = Cell.number(1234.56, format: CellFormat.currency())
        
        sheet.setCell("A1", cell: headerCell)
        sheet.setCell("B1", cell: currencyCell)
        
        let retrievedHeader = sheet.getCellWithFormat("A1")
        let retrievedCurrency = sheet.getCellWithFormat("B1")
        
        XCTAssertNotNil(retrievedHeader)
        XCTAssertNotNil(retrievedCurrency)
        XCTAssertEqual(retrievedHeader?.value, .string("Header"))
        XCTAssertEqual(retrievedCurrency?.value, .number(1234.56))
    }
}
