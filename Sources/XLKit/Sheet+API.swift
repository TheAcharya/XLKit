//
//  Sheet+API.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import CoreGraphics
import XLKitCore
import XLKitImages
import XLKitFormatters

public extension Sheet {
    /// Sets a row of cells with values
    @discardableResult
    func setRow(_ row: Int, values: [CellValue]) -> Self {
        for (index, value) in values.enumerated() {
            let column = index + 1
            setCell(row: row, column: column, value: value)
        }
        return self
    }
    /// Sets a row of cells with strings
    @discardableResult
    func setRow(_ row: Int, strings: [String]) -> Self {
        let values = strings.map { CellValue.string($0) }
        return setRow(row, values: values)
    }
    /// Sets a row of cells with numbers
    @discardableResult
    func setRow(_ row: Int, numbers: [Double]) -> Self {
        let values = numbers.map { CellValue.number($0) }
        return setRow(row, values: values)
    }
    /// Sets a row of cells with integers
    @discardableResult
    func setRow(_ row: Int, integers: [Int]) -> Self {
        let values = integers.map { CellValue.integer($0) }
        return setRow(row, values: values)
    }
    /// Sets a column of cells with values
    @discardableResult
    func setColumn(_ column: Int, values: [CellValue]) -> Self {
        for (index, value) in values.enumerated() {
            let row = index + 1
            setCell(row: row, column: column, value: value)
        }
        return self
    }
    /// Sets a column of cells with strings
    @discardableResult
    func setColumn(_ column: Int, strings: [String]) -> Self {
        let values = strings.map { CellValue.string($0) }
        return setColumn(column, values: values)
    }
    /// Sets a column of cells with numbers
    @discardableResult
    func setColumn(_ column: Int, numbers: [Double]) -> Self {
        let values = numbers.map { CellValue.number($0) }
        return setColumn(column, values: values)
    }
    /// Sets a column of cells with integers
    @discardableResult
    func setColumn(_ column: Int, integers: [Int]) -> Self {
        let values = integers.map { CellValue.integer($0) }
        return setColumn(column, values: values)
    }
    // MARK: - CSV/TSV Export Methods
    /// Exports the sheet to CSV format
    func exportToCSV(separator: String = ",") -> String {
        return CSVUtils.exportToCSV(sheet: self, separator: separator)
    }
    /// Exports the sheet to TSV format
    func exportToTSV() -> String {
        return CSVUtils.exportToTSV(sheet: self)
    }
    // MARK: - Image Embedding Methods
    /// Embeds an image with automatic sizing and aspect ratio preservation
    @MainActor
    @discardableResult
    func embedImageAutoSized(
        _ data: Data,
        at coordinate: String,
        of workbook: Workbook,
        format: ImageFormat? = nil,
        maxCellWidth: CGFloat = 600,
        maxCellHeight: CGFloat = 400,
        scale: CGFloat = 0.5
    ) async throws -> Bool {
        let detected = await ImageUtils.detectImageFormat(from: data)
        let detectedFormat = format ?? detected
        guard let imageFormat = detectedFormat else { return false }
        guard let originalSize = await ImageUtils.getImageSize(from: data, format: imageFormat) else { return false }

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
        addImage(excelImage, at: coordinate)
        workbook.addImage(excelImage)

        // Set column width and row height to fit the scaled image exactly (Excel logic)
        if let cellCoord = CellCoordinate(excelAddress: coordinate) {
            let excelCol = cellCoord.column
            let excelRow = cellCoord.row
    
            // Use the new ImageSizingUtils methods for consistent calculations
            // These match the formulas used in the drawing XML generation
            let colWidth = ImageSizingUtils.excelColumnWidth(forPixelWidth: displaySize.width)
            let rowHeightInPoints = ImageSizingUtils.excelRowHeight(forPixelHeight: displaySize.height)
            
            setColumnWidth(excelCol, width: colWidth)
            setRowHeight(excelRow, height: rowHeightInPoints)
        }
        return true
    }
    /// Embeds an image with scaling
    @MainActor
    @discardableResult
    func embedImage(
        _ data: Data,
        at coordinate: String,
        of workbook: Workbook,
        scale: CGFloat = 0.5,
        maxWidth: CGFloat = 600,
        maxHeight: CGFloat = 400
    ) async throws -> Bool {
        // Calculate scaled max dimensions
        let scaledMaxWidth = maxWidth * scale
        let scaledMaxHeight = maxHeight * scale
        
        return try await embedImageAutoSized(
            data,
            at: coordinate,
            of: workbook,
            format: nil,
            maxCellWidth: scaledMaxWidth,
            maxCellHeight: scaledMaxHeight,
            scale: scale
        )
    }
    /// Embeds an image from a file URL
    @MainActor
    @discardableResult
    func embedImage(from url: URL, at coordinate: String, displaySize: CGSize? = nil) async throws -> Bool {
        guard let image = try await ImageUtils.createExcelImage(from: url, displaySize: displaySize) else { return false }
        addImage(image, at: coordinate)
        return true
    }
    /// Embeds an image from a file path
    @MainActor
    @discardableResult
    func embedImage(from path: String, at coordinate: String, displaySize: CGSize? = nil) async throws -> Bool {
        let url = URL(fileURLWithPath: path)
        return try await embedImage(from: url, at: coordinate, displaySize: displaySize)
    }
    // MARK: - Utility Methods
    var allCells: [String: CellValue] {
        return Dictionary(uniqueKeysWithValues: cells.map { ($0.key, $0.value) })
    }
    var allFormattedCells: [String: Cell] {
        var result: [String: Cell] = [:]
        for (coordinate, value) in cells {
            let format = cellFormats[coordinate]
            result[coordinate] = Cell(value, format: format)
        }
        return result
    }
    var isEmpty: Bool {
        return cells.isEmpty && images.isEmpty && mergedRanges.isEmpty
    }
    var cellCount: Int {
        return cells.count
    }
    var imageCount: Int {
        return images.count
    }
} 