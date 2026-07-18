//
//  ImageTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import CoreGraphics
import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct ImageTests {
    
    @Test func testWorkbookImageManagement() {
        let workbook = Workbook()
        
        let imageData = Data()
        let image = ExcelImage(
            id: "test-image",
            data: imageData,
            format: .png,
            originalSize: CGSize(width: 100, height: 100)
        )
        
        workbook.addImage(image)
        #expect(workbook.imageCount == 1)
        
        let retrievedImage = workbook.getImage(withId: "test-image")
        #expect(retrievedImage != nil)
        #expect(retrievedImage?.id == "test-image")
        
        workbook.removeImage(withId: "test-image")
        #expect(workbook.imageCount == 0)
    }
    
    @Test func testSheetImageManagement() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let imageData = Data()
        let image = ExcelImage(
            id: "test-image",
            data: imageData,
            format: .png,
            originalSize: CGSize(width: 100, height: 100)
        )
        
        sheet.addImage(image, at: "A1")
        #expect(sheet.hasImage(at: "A1"))
        
        let retrievedImage = sheet.getImage(at: "A1")
        #expect(retrievedImage != nil)
        #expect(retrievedImage?.id == "test-image")
        
        sheet.removeImage(at: "A1")
        #expect(!(sheet.hasImage(at: "A1")))
    }
}
