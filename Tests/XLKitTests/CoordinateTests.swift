//
//  CoordinateTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct CoordinateTests {
    
    @Test func testCellCoordinate() {
        // Test coordinate creation
        let coord1 = CellCoordinate(row: 1, column: 1)
        #expect(coord1.excelAddress == "A1")
        
        let coord2 = CellCoordinate(row: 2, column: 2)
        #expect(coord2.excelAddress == "B2")
        
        // Test coordinate from Excel address
        let coord3 = CellCoordinate(excelAddress: "A1")
        #expect(coord3 != nil)
        #expect(coord3?.row == 1)
        #expect(coord3?.column == 1)
        
        let coord4 = CellCoordinate(excelAddress: "B2")
        #expect(coord4 != nil)
        #expect(coord4?.row == 2)
        #expect(coord4?.column == 2)
        
        // Test invalid address
        #expect(CellCoordinate(excelAddress: "Invalid") == nil)
        
        // Test lowercase address (normalized)
        let coord5 = CellCoordinate(excelAddress: "a1")
        #expect(coord5 != nil)
        #expect(coord5?.row == 1)
        #expect(coord5?.column == 1)
        let coord6 = CellCoordinate(excelAddress: "aa10")
        #expect(coord6 != nil)
        #expect(coord6?.row == 10)
        #expect(coord6?.column == 27)
    }
    
    @Test func testCellRange() {
        // Test range creation
        let start = CellCoordinate(row: 1, column: 1)
        let end = CellCoordinate(row: 3, column: 3)
        let range = CellRange(start: start, end: end)
        
        #expect(range.excelRange == "A1:C3")
        #expect(range.coordinates.count == 9) // 3x3 grid
        
        // Test range from Excel notation
        let range2 = CellRange(excelRange: "A1:B2")
        #expect(range2 != nil)
        #expect(range2?.excelRange == "A1:B2")
        #expect(range2?.coordinates.count == 4) // 2x2 grid
        
        // Test invalid range
        #expect(CellRange(excelRange: "Invalid") == nil)
    }
}
