//
//  CoreTypes.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import CoreGraphics
import CryptoKit

/// Core types and data structures for XLKit
/// 
/// This file contains the fundamental types used throughout XLKit:
/// - Workbook: Main container for Excel files with multiple sheets
/// - Sheet: Individual worksheet with cells, formatting, and images
/// - Cell, CellValue, CellFormat: Cell data and styling
/// - CellCoordinate, CellRange: Excel coordinate system
/// - ExcelImage, ImageFormat: Image embedding support
/// - XLKitError: Error handling for all operations
/// - CoreUtils: Utility functions for Excel operations

// MARK: - Workbook

/// Represents an Excel workbook containing multiple sheets and images.
/// 
/// A workbook is the top-level container for Excel files, managing:
/// - Multiple worksheets (Sheet objects)
/// - Embedded images across all sheets
/// - File operations (save, load)
/// - CSV/TSV import/export functionality
/// 
/// ## Usage
/// ```swift
/// let workbook = Workbook()
/// let sheet = workbook.addSheet(name: "Data")
/// 
/// // Add images to workbook
/// let image = ExcelImage(id: "img1", data: imageData, format: .png, originalSize: size)
/// workbook.addImage(image)
/// 
/// // Get workbook information
/// let imageCount = workbook.imageCount
/// let allImages = workbook.getImages()
/// let pngImages = workbook.getImages(withFormat: .png)
/// ```
public final class Workbook {
    private var sheets: [Sheet] = []
    private var nextSheetId: Int
    private var images: [ExcelImage] = []
    
    /// Creates a new empty workbook
    public init() {
        self.nextSheetId = 1
    }
    
    /// Adds a new sheet to the workbook
    /// - Parameter name: The name of the sheet (must be unique)
    /// - Returns: The newly created Sheet object
    @discardableResult
    public func addSheet(name: String) -> Sheet {
        let sheet = Sheet(name: name, id: nextSheetId)
        sheets.append(sheet)
        nextSheetId += 1
        return sheet
    }
    
    /// Gets all sheets in the workbook
    /// - Returns: Array of all Sheet objects
    public func getSheets() -> [Sheet] {
        sheets
    }
    
    /// Gets a sheet by name
    /// - Parameter name: The name of the sheet to find
    /// - Returns: The Sheet object if found, nil otherwise
    public func getSheet(name: String) -> Sheet? {
        sheets.first { $0.name == name }
    }
    
    /// Removes a sheet by name
    /// - Parameter name: The name of the sheet to remove
    public func removeSheet(name: String) {
        sheets.removeAll { $0.name == name }
    }
    
    // MARK: - Image Management
    
    /// Adds an image to the workbook
    /// - Parameter image: The ExcelImage to add
    public func addImage(_ image: ExcelImage) {
        images.append(image)
    }
    
    /// Gets all images in the workbook
    /// - Returns: Array of all ExcelImage objects
    public func getImages() -> [ExcelImage] {
        images
    }
    
    /// Gets an image by ID
    /// - Parameter id: The unique identifier of the image
    /// - Returns: The ExcelImage if found, nil otherwise
    public func getImage(withId id: String) -> ExcelImage? {
        return images.first { $0.id == id }
    }
    
    /// Removes an image by ID
    /// - Parameter id: The unique identifier of the image to remove
    public func removeImage(withId id: String) {
        images.removeAll { $0.id == id }
    }
    
    /// Gets images by format
    /// - Parameter format: The ImageFormat to filter by
    /// - Returns: Array of ExcelImage objects with the specified format
    public func getImages(withFormat format: ImageFormat) -> [ExcelImage] {
        return images.filter { $0.format == format }
    }
    
    /// Clears all images from the workbook
    public func clearImages() {
        images.removeAll()
    }
    
    /// Gets the number of images in the workbook
    public var imageCount: Int {
        return images.count
    }
    
    // MARK: - File Operations
    
    // File operations (save, load) are implemented in Workbook+API.swift to avoid circular dependencies
    
    // MARK: - CSV/TSV Operations
    
    // CSV/TSV operations are implemented in Workbook+API.swift to avoid circular dependencies
}

/// Represents a worksheet in an Excel workbook
/// 
/// A sheet contains cells, formatting, images, and layout information:
/// - Cell data with various value types (string, number, date, etc.)
/// - Cell formatting and styling
/// - Merged cell ranges
/// - Column widths and row heights
/// - Embedded images positioned at specific coordinates
/// 
/// ## Usage
/// ```swift
/// let sheet = workbook.addSheet(name: "Data")
/// 
/// // Set cell values with convenience methods (recommended)
/// sheet.setCell("A1", string: "Header", format: CellFormat.header())
/// sheet.setCell("B1", number: 1234.56, format: CellFormat.currency())
/// sheet.setCell("C1", integer: 100, format: CellFormat.number())
/// 
/// // Set ranges
/// sheet.setRange("A2:C5", string: "Data", format: CellFormat.bordered())
/// 
/// // Merge cells
/// sheet.mergeCells("A1:C1")
/// 
/// // Add images
/// let image = ExcelImage(id: "img1", data: imageData, format: .png, originalSize: size)
/// sheet.addImage(image, at: "D1")
/// ```
public final class Sheet: Equatable {
    /// The name of the sheet
    public let name: String
    /// The unique identifier of the sheet within the workbook
    public let id: Int
    /// Dictionary mapping cell coordinates to their values
    public var cells: [String: CellValue] = [:]
    /// Array of merged cell ranges
    public var mergedRanges: [CellRange] = []
    /// Dictionary mapping column numbers to their widths in pixels
    public var columnWidths: [Int: Double] = [:]
    /// Dictionary mapping row numbers to their heights in pixels
    public var rowHeights: [Int: Double] = [:]
    /// Dictionary mapping cell coordinates to embedded images
    public var images: [String: ExcelImage] = [:]
    /// Dictionary mapping cell coordinates to their formatting
    public var cellFormats: [String: CellFormat] = [:]
    
    public init(name: String, id: Int) {
        self.name = name
        self.id = id
    }
    
    // MARK: - Cell Operations
    
    /// Sets a cell value using the basic CellValue enum
    /// - Parameters:
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    ///   - value: The cell value to set
    /// - Returns: Self for method chaining
    @discardableResult
    public func setCell(_ coordinate: String, value: CellValue) -> Self {
        cells[coordinate.uppercased()] = value
        return self
    }
    
    /// Gets a cell value
    /// - Parameter coordinate: Excel-style coordinate (e.g., "A1", "B2")
    /// - Returns: The cell value if it exists, nil otherwise
    public func getCell(_ coordinate: String) -> CellValue? {
        cells[coordinate.uppercased()]
    }
    
    /// Gets a cell with its formatting information
    /// - Parameter coordinate: Excel-style coordinate (e.g., "A1", "B2")
    /// - Returns: A Cell object containing both value and format, nil if cell doesn't exist
    public func getCellWithFormat(_ coordinate: String) -> Cell? {
        guard let value = cells[coordinate.uppercased()] else { return nil }
        let format = cellFormats[coordinate.uppercased()]
        return Cell(value, format: format)
    }
    
    /// Gets the formatting for a specific cell
    /// - Parameter coordinate: Excel-style coordinate (e.g., "A1", "B2")
    /// - Returns: The cell formatting if it exists, nil otherwise
    public func getCellFormat(_ coordinate: String) -> CellFormat? {
        cellFormats[coordinate.uppercased()]
    }
    
    /// Sets a cell value by row and column
    @discardableResult
    public func setCell(row: Int, column: Int, value: CellValue) -> Self {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        setCell(coordinate, value: value)
        return self
    }
    
    /// Sets a cell with formatting
    @discardableResult
    public func setCell(_ coordinate: String, cell: Cell) -> Self {
        cells[coordinate.uppercased()] = cell.value
        if let format = cell.format {
            cellFormats[coordinate.uppercased()] = format
        }
        return self
    }
    
    /// Sets a cell with formatting by row and column
    @discardableResult
    public func setCell(row: Int, column: Int, cell: Cell) -> Self {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        setCell(coordinate, cell: cell)
        return self
    }
    
    // MARK: - Convenience Cell Methods
    
    /// Sets a cell with a string value and optional formatting
    /// - Parameters:
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    ///   - string: The string value to set
    ///   - format: Optional formatting to apply to the cell
    /// - Returns: Self for method chaining
    @discardableResult
    public func setCell(_ coordinate: String, string: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.string(string, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with a number value and optional formatting
    /// - Parameters:
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    ///   - number: The double value to set
    ///   - format: Optional formatting to apply to the cell
    /// - Returns: Self for method chaining
    @discardableResult
    public func setCell(_ coordinate: String, number: Double, format: CellFormat? = nil) -> Self {
        let cell = Cell.number(number, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with an integer value and optional formatting
    /// - Parameters:
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    ///   - integer: The integer value to set
    ///   - format: Optional formatting to apply to the cell
    /// - Returns: Self for method chaining
    @discardableResult
    public func setCell(_ coordinate: String, integer: Int, format: CellFormat? = nil) -> Self {
        let cell = Cell.integer(integer, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with a boolean value and optional formatting
    /// - Parameters:
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    ///   - boolean: The boolean value to set
    ///   - format: Optional formatting to apply to the cell
    /// - Returns: Self for method chaining
    @discardableResult
    public func setCell(_ coordinate: String, boolean: Bool, format: CellFormat? = nil) -> Self {
        let cell = Cell.boolean(boolean, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with a date value and optional formatting
    /// - Parameters:
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    ///   - date: The date value to set
    ///   - format: Optional formatting to apply to the cell
    /// - Returns: Self for method chaining
    @discardableResult
    public func setCell(_ coordinate: String, date: Date, format: CellFormat? = nil) -> Self {
        let cell = Cell.date(date, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with a formula and optional formatting
    /// - Parameters:
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    ///   - formula: The Excel formula to set (e.g., "=A1+B1", "=SUM(A1:A10)")
    ///   - format: Optional formatting to apply to the cell
    /// - Returns: Self for method chaining
    @discardableResult
    public func setCell(_ coordinate: String, formula: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.formula(formula, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    // MARK: - Range Operations
    
    /// Sets a range of cells with the same value and formatting
    /// - Parameters:
    ///   - range: Excel-style range notation (e.g., "A1:C3", "B2:D5")
    ///   - cell: The Cell object containing value and formatting
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRange(_ range: String, cell: Cell) -> Self {
        guard let cellRange = CellRange(excelRange: range) else { return self }
        for coordinate in cellRange.coordinates {
            setCell(coordinate.excelAddress, cell: cell)
        }
        return self
    }
    
    /// Sets a range of cells with a string value and optional formatting
    /// - Parameters:
    ///   - range: Excel-style range notation (e.g., "A1:C3", "B2:D5")
    ///   - string: The string value to set in all cells in the range
    ///   - format: Optional formatting to apply to all cells in the range
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRange(_ range: String, string: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.string(string, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with a number value and optional formatting
    /// - Parameters:
    ///   - range: Excel-style range notation (e.g., "A1:C3", "B2:D5")
    ///   - number: The double value to set in all cells in the range
    ///   - format: Optional formatting to apply to all cells in the range
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRange(_ range: String, number: Double, format: CellFormat? = nil) -> Self {
        let cell = Cell.number(number, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with an integer value and optional formatting
    /// - Parameters:
    ///   - range: Excel-style range notation (e.g., "A1:C3", "B2:D5")
    ///   - integer: The integer value to set in all cells in the range
    ///   - format: Optional formatting to apply to all cells in the range
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRange(_ range: String, integer: Int, format: CellFormat? = nil) -> Self {
        let cell = Cell.integer(integer, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with a boolean value and optional formatting
    /// - Parameters:
    ///   - range: Excel-style range notation (e.g., "A1:C3", "B2:D5")
    ///   - boolean: The boolean value to set in all cells in the range
    ///   - format: Optional formatting to apply to all cells in the range
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRange(_ range: String, boolean: Bool, format: CellFormat? = nil) -> Self {
        let cell = Cell.boolean(boolean, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with a date value and optional formatting
    /// - Parameters:
    ///   - range: Excel-style range notation (e.g., "A1:C3", "B2:D5")
    ///   - date: The date value to set in all cells in the range
    ///   - format: Optional formatting to apply to all cells in the range
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRange(_ range: String, date: Date, format: CellFormat? = nil) -> Self {
        let cell = Cell.date(date, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with a formula and optional formatting
    /// - Parameters:
    ///   - range: Excel-style range notation (e.g., "A1:C3", "B2:D5")
    ///   - formula: The Excel formula to set in all cells in the range
    ///   - format: Optional formatting to apply to all cells in the range
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRange(_ range: String, formula: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.formula(formula, format: format)
        return setRange(range, cell: cell)
    }
    
    // MARK: - Cell Range Management
    
    /// Merges a range of cells into a single cell
    /// - Parameter range: Excel-style range notation (e.g., "A1:C1", "B2:D5")
    /// - Returns: Self for method chaining
    @discardableResult
    public func mergeCells(_ range: String) -> Self {
        guard let cellRange = CellRange(excelRange: range) else { return self }
        mergedRanges.append(cellRange)
        return self
    }
    
    /// Gets all cell coordinates that have values
    /// - Returns: Array of Excel-style coordinates (e.g., ["A1", "B2", "C3"])
    public func getUsedCells() -> [String] {
        Array(cells.keys).sorted()
    }
    
    /// Gets all merged cell ranges
    /// - Returns: Array of CellRange objects representing merged areas
    public func getMergedRanges() -> [CellRange] {
        mergedRanges
    }
    
    // MARK: - Column and Row Sizing
    
    /// Sets the width of a column in pixels
    /// - Parameters:
    ///   - column: Column number (1-based, where 1 = column A)
    ///   - width: Width in pixels
    /// - Returns: Self for method chaining
    @discardableResult
    public func setColumnWidth(_ column: Int, width: Double) -> Self {
        columnWidths[column] = width
        return self
    }
    
    /// Gets the width of a column in pixels
    /// - Parameter column: Column number (1-based, where 1 = column A)
    /// - Returns: Column width in pixels if set, nil otherwise
    public func getColumnWidth(_ column: Int) -> Double? {
        columnWidths[column]
    }
    
    /// Gets all column widths
    /// - Returns: Dictionary mapping column numbers to their widths in pixels
    public func getColumnWidths() -> [Int: Double] {
        columnWidths
    }
    
    /// Sets the height of a row in pixels
    /// - Parameters:
    ///   - row: Row number (1-based)
    ///   - height: Height in pixels
    /// - Returns: Self for method chaining
    @discardableResult
    public func setRowHeight(_ row: Int, height: Double) -> Self {
        rowHeights[row] = height
        return self
    }
    
    /// Gets the height of a row in pixels
    /// - Parameter row: Row number (1-based)
    /// - Returns: Row height in pixels if set, nil otherwise
    public func getRowHeight(_ row: Int) -> Double? {
        rowHeights[row]
    }
    
    /// Gets all row heights
    /// - Returns: Dictionary mapping row numbers to their heights in pixels
    public func getRowHeights() -> [Int: Double] {
        rowHeights
    }
    
    // MARK: - Image Management
    
    /// Auto-sizes a column to fit an image at a specific coordinate
    /// - Parameters:
    ///   - column: Column number to auto-size (1-based, where 1 = column A)
    ///   - coordinate: Excel-style coordinate where the image is located
    /// - Returns: Self for method chaining
    @discardableResult
    public func autoSizeColumn(_ column: Int, forImageAt coordinate: String) -> Self {
        guard let image = images[coordinate] else { return self }
        let width = image.displaySize?.width ?? image.originalSize.width
        setColumnWidth(column, width: width)
        return self
    }
    
    /// Adds an image to a cell at a specific coordinate
    /// - Parameters:
    ///   - image: The ExcelImage to add
    ///   - coordinate: Excel-style coordinate (e.g., "A1", "B2")
    /// - Returns: Self for method chaining
    @discardableResult
    public func addImage(_ image: ExcelImage, at coordinate: String) -> Self {
        images[coordinate.uppercased()] = image
        return self
    }
    
    /// Adds an image to a cell at a specific row and column
    /// - Parameters:
    ///   - image: The ExcelImage to add
    ///   - row: Row number (1-based)
    ///   - column: Column number (1-based, where 1 = column A)
    /// - Returns: Self for method chaining
    @discardableResult
    public func addImage(_ image: ExcelImage, at row: Int, column: Int) -> Self {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        return addImage(image, at: coordinate)
    }
    
    /// Gets all images in the sheet
    /// - Returns: Dictionary mapping cell coordinates to ExcelImage objects
    public func getImages() -> [String: ExcelImage] {
        images
    }
    
    /// Gets an image at a specific coordinate
    /// - Parameter coordinate: Excel-style coordinate (e.g., "A1", "B2")
    /// - Returns: The ExcelImage if found, nil otherwise
    public func getImage(at coordinate: String) -> ExcelImage? {
        return images[coordinate.uppercased()]
    }
    
    /// Gets an image at a specific row and column
    /// - Parameters:
    ///   - row: Row number (1-based)
    ///   - column: Column number (1-based, where 1 = column A)
    /// - Returns: The ExcelImage if found, nil otherwise
    public func getImage(at row: Int, column: Int) -> ExcelImage? {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        return getImage(at: coordinate)
    }
    
    /// Removes an image from a cell at a specific coordinate
    /// - Parameter coordinate: Excel-style coordinate (e.g., "A1", "B2")
    /// - Returns: Self for method chaining
    @discardableResult
    public func removeImage(at coordinate: String) -> Self {
        images.removeValue(forKey: coordinate.uppercased())
        return self
    }
    
    /// Removes an image from a cell at a specific row and column
    /// - Parameters:
    ///   - row: Row number (1-based)
    ///   - column: Column number (1-based, where 1 = column A)
    /// - Returns: Self for method chaining
    @discardableResult
    public func removeImage(at row: Int, column: Int) -> Self {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        return removeImage(at: coordinate)
    }
    
    /// Checks if a cell has an image at a specific coordinate
    /// - Parameter coordinate: Excel-style coordinate (e.g., "A1", "B2")
    /// - Returns: True if the cell contains an image, false otherwise
    public func hasImage(at coordinate: String) -> Bool {
        return images[coordinate.uppercased()] != nil
    }
    
    /// Checks if a cell has an image at a specific row and column
    /// - Parameters:
    ///   - row: Row number (1-based)
    ///   - column: Column number (1-based, where 1 = column A)
    /// - Returns: True if the cell contains an image, false otherwise
    public func hasImage(at row: Int, column: Int) -> Bool {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        return hasImage(at: coordinate)
    }
    
    // MARK: - Utility Methods
    
    /// Clears all data from the sheet (cells, formatting, images, merged ranges, dimensions)
    public func clear() {
        cells.removeAll()
        mergedRanges.removeAll()
        columnWidths.removeAll()
        rowHeights.removeAll()
        images.removeAll()
        cellFormats.removeAll()
    }
    
    /// Equality comparison for Sheet objects (reference equality)
    public static func == (lhs: Sheet, rhs: Sheet) -> Bool {
        lhs === rhs // Reference equality
    }
}

// MARK: - Cell Coordinate System

/// Represents a cell coordinate in Excel notation
/// 
/// Provides conversion between Excel-style addresses (A1, B2) and row/column numbers.
/// Supports both creation from row/column numbers and parsing from Excel addresses.
/// 
/// ## Usage
/// ```swift
/// // Create from row and column numbers
/// let coord1 = CellCoordinate(row: 1, column: 1) // A1
/// let coord2 = CellCoordinate(row: 2, column: 2) // B2
/// 
/// // Parse from Excel address
/// let coord3 = CellCoordinate(excelAddress: "A1") // row: 1, column: 1
/// let coord4 = CellCoordinate(excelAddress: "B2") // row: 2, column: 2
/// 
/// // Get Excel address
/// print(coord1.excelAddress) // "A1"
/// print(coord2.excelAddress) // "B2"
/// ```
public struct CellCoordinate: Hashable {
    /// Row number (1-based)
    public let row: Int
    /// Column number (1-based, where 1 = column A)
    public let column: Int

    public init(row: Int, column: Int) {
        self.row = row
        self.column = column
    }
    
    /// Excel-style address (e.g., "A1", "B2")
    public var excelAddress: String {
        return CoreUtils.columnLetter(from: column) + String(row)
    }
    
    /// Creates a coordinate from Excel address
    /// - Parameter excelAddress: Excel-style address (e.g., "A1", "B2", "AA10")
    public init?(excelAddress: String) {
        let pattern = "^([A-Z]+)([0-9]+)$"
        guard let regex = try? NSRegularExpression(pattern: pattern),
              let match = regex.firstMatch(in: excelAddress, range: NSRange(excelAddress.startIndex..., in: excelAddress)) else {
            return nil
        }
        
        let columnRange = Range(match.range(at: 1), in: excelAddress)!
        let rowRange = Range(match.range(at: 2), in: excelAddress)!
        
        let columnString = String(excelAddress[columnRange])
        let rowString = String(excelAddress[rowRange])
        
        guard let row = Int(rowString) else { return nil }
        let column = CoreUtils.columnNumber(from: columnString)
        
        self.row = row
        self.column = column
    }
}

/// Represents a range of cells in Excel notation
/// 
/// Provides functionality for working with cell ranges, including creation from Excel range notation
/// and iteration over all coordinates within the range.
/// 
/// ## Usage
/// ```swift
/// // Create from start and end coordinates
/// let start = CellCoordinate(row: 1, column: 1) // A1
/// let end = CellCoordinate(row: 3, column: 3)   // C3
/// let range = CellRange(start: start, end: end) // A1:C3
/// 
/// // Parse from Excel range notation
/// let range2 = CellRange(excelRange: "A1:B5") // A1 to B5
/// 
/// // Get all coordinates in range
/// let coords = range.coordinates // [A1, A2, A3, B1, B2, B3, C1, C2, C3]
/// 
/// // Get Excel range notation
/// print(range.excelRange) // "A1:C3"
/// ```
public struct CellRange: Hashable {
    /// Starting coordinate of the range
    public let start: CellCoordinate
    /// Ending coordinate of the range
    public let end: CellCoordinate
    
    /// Creates a range from start and end coordinates
    /// - Parameters:
    ///   - start: Starting coordinate
    ///   - end: Ending coordinate
    public init(start: CellCoordinate, end: CellCoordinate) {
        self.start = start
        self.end = end
    }
    
    /// Creates a range from Excel range notation
    /// - Parameter excelRange: Excel-style range notation (e.g., "A1:B5", "C3:D10")
    public init?(excelRange: String) {
        let components = excelRange.components(separatedBy: ":")
        guard components.count == 2,
              let start = CellCoordinate(excelAddress: components[0]),
              let end = CellCoordinate(excelAddress: components[1]) else {
            return nil
        }
        self.start = start
        self.end = end
    }
    
    /// All coordinates within this range
    /// - Returns: Array of CellCoordinate objects representing all cells in the range
    public var coordinates: [CellCoordinate] {
        var result: [CellCoordinate] = []
        for row in start.row...end.row {
            for column in start.column...end.column {
                result.append(CellCoordinate(row: row, column: column))
            }
        }
        return result
    }
    
    /// Excel range notation
    /// - Returns: String representation of the range (e.g., "A1:C3")
    public var excelRange: String {
        return "\(start.excelAddress):\(end.excelAddress)"
    }
}

/// Represents a cell value
public enum CellValue: Equatable {
    case string(String)
    case number(Double)
    case integer(Int)
    case boolean(Bool)
    case date(Date)
    case formula(String)
    case empty
    
    /// String representation of the value
    public var stringValue: String {
        switch self {
        case .string(let value):
            return value
        case .number(let value):
            return String(value)
        case .integer(let value):
            return String(value)
        case .boolean(let value):
            return value ? "TRUE" : "FALSE"
        case .date(let value):
            return CoreUtils.formatDate(value)
        case .formula(let value):
            return value
        case .empty:
            return ""
        }
    }
    
    /// Type of the value
    public var type: String {
        switch self {
        case .string: return "string"
        case .number: return "number"
        case .integer: return "integer"
        case .boolean: return "boolean"
        case .date: return "date"
        case .formula: return "formula"
        case .empty: return "empty"
        }
    }
}

/// Represents a cell with value and formatting
public struct Cell {
    public let value: CellValue
    public let format: CellFormat?
    
    public init(_ value: CellValue, format: CellFormat? = nil) {
        self.value = value
        self.format = format
    }
    
    /// Creates a cell with string value
    public static func string(_ value: String, format: CellFormat? = nil) -> Cell {
        return Cell(.string(value), format: format)
    }
    
    /// Creates a cell with number value
    public static func number(_ value: Double, format: CellFormat? = nil) -> Cell {
        return Cell(.number(value), format: format)
    }
    
    /// Creates a cell with integer value
    public static func integer(_ value: Int, format: CellFormat? = nil) -> Cell {
        return Cell(.integer(value), format: format)
    }
    
    /// Creates a cell with boolean value
    public static func boolean(_ value: Bool, format: CellFormat? = nil) -> Cell {
        return Cell(.boolean(value), format: format)
    }
    
    /// Creates a cell with date value
    public static func date(_ value: Date, format: CellFormat? = nil) -> Cell {
        return Cell(.date(value), format: format)
    }
    
    /// Creates a cell with formula
    public static func formula(_ value: String, format: CellFormat? = nil) -> Cell {
        return Cell(.formula(value), format: format)
    }
}

// MARK: - Cell Formatting

/// Font weight for text formatting
public enum FontWeight: String {
    case normal = "normal"
    case bold = "bold"
}

/// Font style for text formatting
public enum FontStyle: String {
    case normal = "normal"
    case italic = "italic"
}

/// Text decoration for formatting
public enum TextDecoration: String {
    case none = "none"
    case underline = "underline"
    case strikethrough = "strikethrough"
    case underlineStrikethrough = "underlineStrikethrough"
}

/// Horizontal alignment for cells
public enum HorizontalAlignment: String {
    case left = "left"
    case center = "center"
    case right = "right"
    case justify = "justify"
    case distributed = "distributed"
}

/// Vertical alignment for cells
public enum VerticalAlignment: String {
    case top = "top"
    case center = "center"
    case bottom = "bottom"
    case justify = "justify"
    case distributed = "distributed"
}

/// Border style for cells
public enum BorderStyle: String {
    case none = "none"
    case thin = "thin"
    case medium = "medium"
    case thick = "thick"
    case double = "double"
    case hair = "hair"
    case dotted = "dotted"
    case dashed = "dashed"
    case dashDot = "dashDot"
    case dashDotDot = "dashDotDot"
    case slantDashDot = "slantDashDot"
}

/// Number format types
public enum NumberFormat: String {
    case general = "General"
    case number = "0"
    case numberWithDecimals = "0.00"
    case percentage = "0%"
    case percentageWithDecimals = "0.00%"
    case currency = "$#,##0"
    case currencyWithDecimals = "$#,##0.00"
    case date = "m/d/yyyy"
    case dateTime = "m/d/yyyy h:mm"
    case time = "h:mm:ss"
    case text = "@"
    case custom = "custom"
    
    /// Creates a custom number format
    public static func custom(_ format: String) -> NumberFormat {
        return .custom
    }
}

/// Represents cell formatting options
public struct CellFormat {
    public var fontName: String?
    public var fontSize: Double?
    public var fontWeight: FontWeight?
    public var fontStyle: FontStyle?
    public var textDecoration: TextDecoration?
    public var fontColor: String? // Hex color like "#FF0000"
    public var backgroundColor: String? // Hex color like "#FFFF00"
    
    public var horizontalAlignment: HorizontalAlignment?
    public var verticalAlignment: VerticalAlignment?
    public var textWrapping: Bool?
    public var textRotation: Int? // Degrees, 0-180
    
    public var numberFormat: NumberFormat?
    public var customNumberFormat: String?
    
    public var borderTop: BorderStyle?
    public var borderBottom: BorderStyle?
    public var borderLeft: BorderStyle?
    public var borderRight: BorderStyle?
    public var borderColor: String? // Hex color
    
    public init() {}
    
    /// Creates a format with basic text styling
    public static func text(
        fontName: String? = nil,
        fontSize: Double? = nil,
        fontWeight: FontWeight? = nil,
        fontStyle: FontStyle? = nil,
        color: String? = nil
    ) -> CellFormat {
        var format = CellFormat()
        format.fontName = fontName
        format.fontSize = fontSize
        format.fontWeight = fontWeight
        format.fontStyle = fontStyle
        format.fontColor = color
        return format
    }
    
    /// Creates a format for headers
    public static func header(
        fontSize: Double = 14.0,
        backgroundColor: String = "#E6E6E6"
    ) -> CellFormat {
        var format = CellFormat()
        format.fontWeight = .bold
        format.fontSize = fontSize
        format.backgroundColor = backgroundColor
        format.horizontalAlignment = .center
        return format
    }
    
    /// Creates a format for currency
    public static func currency(
        format: NumberFormat = .currencyWithDecimals,
        color: String? = nil
    ) -> CellFormat {
        var cellFormat = CellFormat()
        cellFormat.numberFormat = format
        cellFormat.fontColor = color
        return cellFormat
    }
    
    /// Creates a format for percentages
    public static func percentage(
        format: NumberFormat = .percentageWithDecimals
    ) -> CellFormat {
        var cellFormat = CellFormat()
        cellFormat.numberFormat = format
        return cellFormat
    }
    
    /// Creates a format for dates
    public static func date(
        format: NumberFormat = .date
    ) -> CellFormat {
        var cellFormat = CellFormat()
        cellFormat.numberFormat = format
        return cellFormat
    }
    
    /// Creates a format with borders
    public static func bordered(
        style: BorderStyle = .thin,
        color: String? = nil
    ) -> CellFormat {
        var format = CellFormat()
        format.borderTop = style
        format.borderBottom = style
        format.borderLeft = style
        format.borderRight = style
        format.borderColor = color
        return format
    }
}

/// Error types for XLKit operations
public enum XLKitError: Error, LocalizedError {
    case invalidCoordinate(String)
    case invalidRange(String)
    case fileWriteError(String)
    case zipCreationError(String)
    case xmlGenerationError(String)
    case securityError(String)
    case rateLimitExceeded(String)
    case fileSizeLimitExceeded(String)
    case suspiciousFileDetected(String)
    
    public var errorDescription: String? {
        switch self {
        case .invalidCoordinate(let coord):
            return "Invalid coordinate: \(coord)"
        case .invalidRange(let range):
            return "Invalid range: \(range)"
        case .fileWriteError(let message):
            return "File write error: \(message)"
        case .zipCreationError(let message):
            return "ZIP creation error: \(message)"
        case .xmlGenerationError(let message):
            return "XML generation error: \(message)"
        case .securityError(let message):
            return "Security error: \(message)"
        case .rateLimitExceeded(let message):
            return "Rate limit exceeded: \(message)"
        case .fileSizeLimitExceeded(let message):
            return "File size limit exceeded: \(message)"
        case .suspiciousFileDetected(let message):
            return "Suspicious file detected: \(message)"
        }
    }
}

// MARK: - Core Utilities

/// Core utility functions for XLKit
public struct CoreUtils {
    
    // MARK: - Column/Row Conversion
    
    /// Converts column number to Excel letter (1 -> A, 2 -> B, etc.)
    public static func columnLetter(from column: Int) -> String {
        var result = ""
        var col = column
        
        while col > 0 {
            col -= 1
            let char = Character(UnicodeScalar(65 + (col % 26))!)
            result = String(char) + result
            col /= 26
        }
        
        return result
    }
    
    /// Converts Excel column letter to number (A -> 1, B -> 2, etc.)
    public static func columnNumber(from letter: String) -> Int {
        var result = 0
        let upperLetter = letter.uppercased()
        
        for char in upperLetter {
            result = result * 26 + Int(char.asciiValue! - 64)
        }
        
        return result
    }
    
    // MARK: - Date Formatting
    
    /// Formats a date for Excel
    public static func formatDate(_ date: Date) -> String {
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyy-mm-dd"
        return formatter.string(from: date)
    }
    
    /// Converts Excel date number to Date
    public static func dateFromExcelNumber(_ number: Double) -> Date {
        // Excel dates are days since January 1, 1900
        let excelEpoch = Date(timeIntervalSince1970: -2208988800) // 1900-01-01
        let days = number - 2 // Excel has a leap year bug, so subtract 2
        return excelEpoch.addingTimeInterval(days * 24 * 60 * 60)
    }
    
    /// Converts Date to Excel date number
    public static func excelNumberFromDate(_ date: Date) -> Double {
        let excelEpoch = Date(timeIntervalSince1970: -2208988800) // 1900-01-01
        let days = date.timeIntervalSince(excelEpoch) / (24 * 60 * 60)
        return days + 2 // Add 2 to account for Excel's leap year bug
    }
    
    // MARK: - XML Generation
    
    /// Escapes XML special characters
    public static func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
    
    /// Generates XML header
    public static func xmlHeader() -> String {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    }
    
    // MARK: - Security Utilities
    
    /// Maximum allowed file size for images (50MB)
    public static let maxImageFileSize = 50 * 1024 * 1024
    
    /// Maximum allowed file size for Excel files (100MB)
    public static let maxExcelFileSize = 100 * 1024 * 1024
    
    /// Allowed directories for file operations
    public static let allowedDirectories: [URL] = {
        var directories = [FileManager.default.temporaryDirectory]
        
        #if os(macOS)
        directories.append(FileManager.default.homeDirectoryForCurrentUser)
        #endif
        
        return directories
    }()
    
    /// Validates file path for security
    @MainActor
    public static func validateFilePath(_ path: String) throws {
        // Skip validation if file path restrictions are disabled
        guard SecurityManager.enableFilePathRestrictions else {
            return
        }
        
        // Prevent path traversal attacks
        guard !path.contains("..") else {
            throw XLKitError.securityError("Path traversal not allowed")
        }
        
        // Ensure path is within allowed directories
        let fileURL = URL(fileURLWithPath: path)
        
        guard allowedDirectories.contains(where: { fileURL.path.hasPrefix($0.path) }) else {
            throw XLKitError.securityError("File path outside allowed directories")
        }
    }
    
    /// Validates file size
    public static func validateFileSize(_ data: Data, maxSize: Int) throws {
        guard data.count <= maxSize else {
            throw XLKitError.fileSizeLimitExceeded("File size \(data.count) exceeds limit \(maxSize)")
        }
    }
    
    /// Validates image file size
    public static func validateImageFileSize(_ data: Data) throws {
        try validateFileSize(data, maxSize: maxImageFileSize)
    }
    
    /// Validates Excel file size
    public static func validateExcelFileSize(_ data: Data) throws {
        try validateFileSize(data, maxSize: maxExcelFileSize)
    }
    
    /// Generates SHA-256 checksum for data
    public static func generateChecksum(_ data: Data) -> String {
        let hash = SHA256.hash(data: data)
        return hash.compactMap { String(format: "%02x", $0) }.joined()
    }
    
    /// Generates SHA-256 checksum for file
    public static func generateFileChecksum(_ fileURL: URL) throws -> String {
        let data = try Data(contentsOf: fileURL)
        return generateChecksum(data)
    }
}

/// Supported image formats for Excel embedding
public enum ImageFormat: String, CaseIterable {
    case gif = "gif"
    case png = "png"
    case jpeg = "jpeg"
    case jpg = "jpg"
    
    /// MIME type for the image format
    public var mimeType: String {
        switch self {
        case .gif: return "image/gif"
        case .png: return "image/png"
        case .jpeg, .jpg: return "image/jpeg"
        }
    }
    
    /// Excel content type for the image format
    public var excelContentType: String {
        switch self {
        case .gif: return "image/gif"
        case .png: return "image/png"
        case .jpeg, .jpg: return "image/jpeg"
        }
    }
}

/// Represents an image to be embedded in Excel
public struct ExcelImage {
    public let id: String
    public let data: Data
    public let format: ImageFormat
    public let originalSize: CGSize
    public let displaySize: CGSize?
    
    public init(id: String, data: Data, format: ImageFormat, originalSize: CGSize, displaySize: CGSize? = nil) {
        self.id = id
        self.data = data
        self.format = format
        self.originalSize = originalSize
        self.displaySize = displaySize
    }
} 