//
//  XLKit.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

@_exported import XLKitCore
@_exported import XLKitFormatters
@_exported import XLKitImages
@_exported import XLKitXLSX

/// XLKit main API: re-exports all public APIs from submodules for easy use.

import Foundation

public struct XLKit {
    
    // MARK: - Workbook Creation
    
    /// Creates a new workbook
    public static func createWorkbook() -> Workbook {
        return Workbook()
    }
    
    /// Creates a workbook from CSV data
    public static func createWorkbookFromCSV(
        csvData: String, 
        sheetName: String = "Sheet1", 
        separator: String = ",", 
        hasHeader: Bool = false
    ) -> Workbook {
        return CSVUtils.createWorkbookFromCSV(
            csvData: csvData, 
            sheetName: sheetName, 
            separator: separator, 
            hasHeader: hasHeader
        )
    }
    
    /// Creates a workbook from TSV data
    public static func createWorkbookFromTSV(
        tsvData: String, 
        sheetName: String = "Sheet1", 
        hasHeader: Bool = false
    ) -> Workbook {
        return CSVUtils.createWorkbookFromTSV(
            tsvData: tsvData, 
            sheetName: sheetName, 
            hasHeader: hasHeader
        )
    }
    
    // MARK: - File Operations
    
    /// Saves a workbook to a file synchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) throws {
        try XLSXEngine.generateXLSX(workbook: workbook, to: url)
    }
    
    /// Saves a workbook to a file asynchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) async throws {
        try XLSXEngine.generateXLSX(workbook: workbook, to: url)
    }
    
    // MARK: - CSV/TSV Export Methods
    
    /// Exports a sheet to CSV format
    public static func exportSheetToCSV(sheet: Sheet, separator: String = ",") -> String {
        return CSVUtils.exportToCSV(sheet: sheet, separator: separator)
    }
    
    /// Exports a sheet to TSV format
    public static func exportSheetToTSV(sheet: Sheet) -> String {
        return CSVUtils.exportToTSV(sheet: sheet)
    }
    
    // MARK: - CSV/TSV Import Methods
    
    /// Imports CSV data into a sheet
    public static func importCSVIntoSheet(
        sheet: Sheet, 
        csvData: String, 
        separator: String = ",", 
        hasHeader: Bool = false
    ) {
        CSVUtils.importFromCSV(
            sheet: sheet, 
            csvData: csvData, 
            separator: separator, 
            hasHeader: hasHeader
        )
    }
    
    /// Imports TSV data into a sheet
    public static func importTSVIntoSheet(
        sheet: Sheet, 
        tsvData: String, 
        hasHeader: Bool = false
    ) {
        CSVUtils.importFromTSV(
            sheet: sheet, 
            tsvData: tsvData, 
            hasHeader: hasHeader
        )
    }
    
    // MARK: - Image Embedding Convenience Methods
    
    /// Embeds an image from data into a sheet at the specified coordinate
    @discardableResult
    public static func embedImage(
        _ data: Data,
        at coordinate: String,
        in sheet: Sheet,
        format: ImageFormat,
        displaySize: CGSize? = nil
    ) -> Bool {
        return sheet.addImage(data, at: coordinate, format: format, displaySize: displaySize)
    }
    
    /// Embeds an image from a file URL into a sheet at the specified coordinate
    @discardableResult
    public static func embedImage(
        from url: URL,
        at coordinate: String,
        in sheet: Sheet,
        displaySize: CGSize? = nil
    ) throws -> Bool {
        return try sheet.addImage(from: url, at: coordinate, displaySize: displaySize)
    }
    
    /// Embeds an image from a file path into a sheet at the specified coordinate
    @discardableResult
    public static func embedImage(
        from path: String,
        at coordinate: String,
        in sheet: Sheet,
        displaySize: CGSize? = nil
    ) throws -> Bool {
        let url = URL(fileURLWithPath: path)
        return try embedImage(from: url, at: coordinate, in: sheet, displaySize: displaySize)
    }
    
    /// Embeds an image from data into a sheet at the specified coordinate, maintaining aspect ratio and auto-sizing the cell
    @discardableResult
    public static func embedImageAutoSized(
        _ data: Data,
        at coordinate: String,
        in sheet: Sheet,
        of workbook: Workbook,
        format: ImageFormat? = nil,
        maxCellWidth: CGFloat = 600,
        maxCellHeight: CGFloat = 400
    ) -> Bool {
        let detectedFormat = format ?? ImageUtils.detectImageFormat(from: data)
        guard let imageFormat = detectedFormat else { return false }
        guard let originalSize = ImageUtils.getImageSize(from: data, format: imageFormat) else { return false }
        
        // Dynamic sizing logic based on original image dimensions
        let originalWidth = originalSize.width
        let originalHeight = originalSize.height
        
        // Define reasonable display limits for Excel (match Notion export logic)
        let maxDisplayWidth: CGFloat = maxCellWidth  // Maximum display width in pixels
        let maxDisplayHeight: CGFloat = maxCellHeight // Maximum display height in pixels
        let minDisplayWidth: CGFloat = 100  // Minimum display width in pixels
        let minDisplayHeight: CGFloat = 60  // Minimum display height in pixels
        
        // Calculate scale to fit within max bounds while maintaining aspect ratio
        let widthScale = maxDisplayWidth / originalWidth
        let heightScale = maxDisplayHeight / originalHeight
        let maxScale = min(widthScale, heightScale, 1.0) // Don't scale up, only down
        
        // Calculate initial display size
        var displayWidth = originalWidth * maxScale
        var displayHeight = originalHeight * maxScale
        
        // Ensure minimum size for very small images
        if displayWidth < minDisplayWidth || displayHeight < minDisplayHeight {
            let minWidthScale = minDisplayWidth / originalWidth
            let minHeightScale = minDisplayHeight / originalHeight
            let minScale = max(minWidthScale, minHeightScale)
            displayWidth = originalWidth * minScale
            displayHeight = originalHeight * minScale
        }
        
        let displaySize = CGSize(width: displayWidth, height: displayHeight)
        
        // Create ExcelImage
        let excelImage = ExcelImage(
            id: "image_\(UUID().uuidString)",
            data: data,
            format: imageFormat,
            originalSize: originalSize,
            displaySize: displaySize
        )
        sheet.addImage(excelImage, at: coordinate)
        workbook.addImage(excelImage)
        
        // Set column width and row height to fit the image exactly (Excel logic)
        if let cellCoord = CellCoordinate(excelAddress: coordinate) {
            let excelCol = cellCoord.column
            let excelRow = cellCoord.row
            // Use precise Excel formulas for pixel-perfect sizing
            let roundedWidth = round(displayWidth)
            let roundedHeight = round(displayHeight)
            // Excel column width formula: width = (pixels - 5) / 7
            let colWidth = max((roundedWidth - 5.0) / 7.0, 0.0)
            // Excel row height formula: height = pixels * 72 / 96
            let rowHeightInPoints = roundedHeight * 72.0 / 96.0
            
            // Set column width and row height to match image dimensions exactly
            sheet.setColumnWidth(excelCol, width: colWidth)
            sheet.setRowHeight(excelRow, height: rowHeightInPoints)
            

        }
        
        return true
    }
    
    /// Creates an ExcelImage from data with automatic format detection
    public static func createImage(
        from data: Data,
        displaySize: CGSize? = nil
    ) -> ExcelImage? {
        guard let format = ImageUtils.detectImageFormat(from: data) else { return nil }
        return ImageUtils.createExcelImage(from: data, format: format, displaySize: displaySize)
    }
    
    /// Creates an ExcelImage from a file URL with automatic format detection
    public static func createImage(
        from url: URL,
        displaySize: CGSize? = nil
    ) throws -> ExcelImage? {
        return try ImageUtils.createExcelImage(from: url, displaySize: displaySize)
    }
    
    /// Creates an ExcelImage from a file path with automatic format detection
    public static func createImage(
        from path: String,
        displaySize: CGSize? = nil
    ) throws -> ExcelImage? {
        let url = URL(fileURLWithPath: path)
        return try createImage(from: url, displaySize: displaySize)
    }
    
    // MARK: - Utility Methods
    
    /// Detects image format from data
    public static func detectImageFormat(from data: Data) -> ImageFormat? {
        return ImageUtils.detectImageFormat(from: data)
    }
    
    /// Gets image size from data
    public static func getImageSize(from data: Data, format: ImageFormat) -> CGSize? {
        return ImageUtils.getImageSize(from: data, format: format)
    }
}

public extension Sheet {
    /// Embeds an image at a coordinate, maintaining aspect ratio and auto-sizing the cell
    @discardableResult
    func embedImageAutoSized(
        _ data: Data,
        at coordinate: String,
        of workbook: Workbook,
        format: ImageFormat? = nil,
        maxCellWidth: CGFloat = 600,
        maxCellHeight: CGFloat = 400
    ) -> Bool {
        XLKit.embedImageAutoSized(
            data,
            at: coordinate,
            in: self,
            of: workbook,
            format: format,
            maxCellWidth: maxCellWidth,
            maxCellHeight: maxCellHeight
        )
    }
}
