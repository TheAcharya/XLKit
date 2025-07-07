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
@preconcurrency import XLKitCore
import XLKitImages // Add this import for ImageSizingUtils

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
    
    /// Embeds an image from data into a sheet at the specified coordinate with automatic sizing and aspect ratio preservation
    /// - Parameter scale: Scaling factor for image size (default: 0.5 = 50% of max bounds for more compact images)
    @discardableResult
    public static func embedImage(
        _ data: Data,
        at coordinate: String,
        in sheet: Sheet,
        of workbook: Workbook,
        scale: CGFloat = 0.5, // Default 50% scaling for more compact images
        maxWidth: CGFloat = 600,
        maxHeight: CGFloat = 400
    ) -> Bool {
        // Calculate scaled max dimensions
        let scaledMaxWidth = maxWidth * scale
        let scaledMaxHeight = maxHeight * scale
        
        return embedImageAutoSized(
            data,
            at: coordinate,
            in: sheet,
            of: workbook,
            format: nil,
            maxCellWidth: scaledMaxWidth,
            maxCellHeight: scaledMaxHeight
        )
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
    /// - Parameter scale: Scaling factor for image size (default: 0.5 = 50% of max bounds for more compact images)
    @discardableResult
    public static func embedImageAutoSized(
        _ data: Data,
        at coordinate: String,
        in sheet: Sheet,
        of workbook: Workbook,
        format: ImageFormat? = nil,
        maxCellWidth: CGFloat = 600,
        maxCellHeight: CGFloat = 400,
        scale: CGFloat = 0.5 // Default 50% scaling for more compact images
    ) -> Bool {
        let detectedFormat = format ?? ImageUtils.detectImageFormat(from: data)
        guard let imageFormat = detectedFormat else { return false }
        guard let originalSize = ImageUtils.getImageSize(from: data, format: imageFormat) else { return false }

        // Use new utility for sizing
        let displaySize = ImageSizingUtils.calculateDisplaySize(
            originalSize: originalSize,
            maxWidth: maxCellWidth,
            maxHeight: maxCellHeight,
            scale: scale
        )

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

        // Set column width and row height to fit the scaled image exactly (Excel logic)
        if let cellCoord = CellCoordinate(excelAddress: coordinate) {
            let excelCol = cellCoord.column
            let excelRow = cellCoord.row
            
            // Use the new ImageSizingUtils methods for consistent calculations
            // These match the formulas used in the drawing XML generation
            let colWidth = ImageSizingUtils.excelColumnWidth(forPixelWidth: displaySize.width)
            let rowHeightInPoints = ImageSizingUtils.excelRowHeight(forPixelHeight: displaySize.height)
            
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
    /// - Parameter scale: Scaling factor for image size (default: 0.5 = 50% of max bounds for more compact images)
    @discardableResult
    func embedImageAutoSized(
        _ data: Data,
        at coordinate: String,
        of workbook: Workbook,
        format: ImageFormat? = nil,
        maxCellWidth: CGFloat = 600,
        maxCellHeight: CGFloat = 400,
        scale: CGFloat = 0.5 // Default 50% scaling for more compact images
    ) -> Bool {
        XLKit.embedImageAutoSized(
            data,
            at: coordinate,
            in: self,
            of: workbook,
            format: format,
            maxCellWidth: maxCellWidth,
            maxCellHeight: maxCellHeight,
            scale: scale
        )
    }
    
    /// Embeds an image at a coordinate with automatic sizing and aspect ratio preservation (simplified API)
    /// - Parameter scale: Scaling factor for image size (default: 0.5 = 50% of max bounds for more compact images)
    @discardableResult
    func embedImage(
        _ data: Data,
        at coordinate: String,
        of workbook: Workbook,
        scale: CGFloat = 0.5, // Default 50% scaling for more compact images
        maxWidth: CGFloat = 600,
        maxHeight: CGFloat = 400
    ) -> Bool {
        XLKit.embedImage(
            data,
            at: coordinate,
            in: self,
            of: workbook,
            scale: scale,
            maxWidth: maxWidth,
            maxHeight: maxHeight
        )
    }
}
