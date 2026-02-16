//
//  ImageTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class ImageTests: XLKitTestBase {
    
    func testWorkbookImageManagement() {
        let workbook = Workbook()
        
        let imageData = Data()
        let image = ExcelImage(
            id: "test-image",
            data: imageData,
            format: .png,
            originalSize: CGSize(width: 100, height: 100)
        )
        
        workbook.addImage(image)
        XCTAssertEqual(workbook.imageCount, 1)
        
        let retrievedImage = workbook.getImage(withId: "test-image")
        XCTAssertNotNil(retrievedImage)
        XCTAssertEqual(retrievedImage?.id, "test-image")
        
        workbook.removeImage(withId: "test-image")
        XCTAssertEqual(workbook.imageCount, 0)
    }
    
    func testSheetImageManagement() {
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
        XCTAssertTrue(sheet.hasImage(at: "A1"))
        
        let retrievedImage = sheet.getImage(at: "A1")
        XCTAssertNotNil(retrievedImage)
        XCTAssertEqual(retrievedImage?.id, "test-image")
        
        sheet.removeImage(at: "A1")
        XCTAssertFalse(sheet.hasImage(at: "A1"))
    }
}
