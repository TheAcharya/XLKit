import Foundation

/// Universal Excel library for creating and manipulating .xlsx files
public struct XLKit {
    
    /// Creates a new workbook
    public static func createWorkbook() -> Workbook {
        Workbook()
    }
    
    /// Saves a workbook to a file asynchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) async throws {
        try await workbook.save(to: url)
    }
    
    /// Saves a workbook to a file synchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) throws {
        try workbook.saveSync(to: url)
    }
}

/// Supported image formats for Excel embedding
public enum ImageFormat: String, CaseIterable {
    case gif = "gif"
    case png = "png"
    case jpeg = "jpeg"
    case jpg = "jpg"
    case bmp = "bmp"
    case tiff = "tiff"
    
    /// MIME type for the image format
    public var mimeType: String {
        switch self {
        case .gif: return "image/gif"
        case .png: return "image/png"
        case .jpeg, .jpg: return "image/jpeg"
        case .bmp: return "image/bmp"
        case .tiff: return "image/tiff"
        }
    }
    
    /// Excel content type for the image format
    public var excelContentType: String {
        switch self {
        case .gif: return "image/gif"
        case .png: return "image/png"
        case .jpeg, .jpg: return "image/jpeg"
        case .bmp: return "image/bmp"
        case .tiff: return "image/tiff"
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
    
    /// Creates an ExcelImage from Data
    public static func from(data: Data, format: ImageFormat, displaySize: CGSize? = nil) -> ExcelImage? {
        guard let size = XLKitUtils.getImageSize(from: data, format: format) else { return nil }
        let id = "image_\(UUID().uuidString)"
        return ExcelImage(id: id, data: data, format: format, originalSize: size, displaySize: displaySize)
    }
    
    /// Creates an ExcelImage from a file URL
    public static func from(url: URL, displaySize: CGSize? = nil) throws -> ExcelImage? {
        let data = try Data(contentsOf: url)
        guard let format = XLKitUtils.detectImageFormat(from: data) else { return nil }
        return from(data: data, format: format, displaySize: displaySize)
    }
}

/// Represents an Excel workbook containing multiple sheets
public final class Workbook {
    private var sheets: [Sheet] = []
    private let nextSheetId: Int
    private var images: [ExcelImage] = []
    
    public init() {
        self.nextSheetId = 1
    }
    
    /// Adds a new sheet to the workbook
    @discardableResult
    public func addSheet(name: String) -> Sheet {
        let sheet = Sheet(name: name, id: nextSheetId)
        sheets.append(sheet)
        return sheet
    }
    
    /// Gets all sheets in the workbook
    public func getSheets() -> [Sheet] {
        sheets
    }
    
    /// Gets a sheet by name
    public func getSheet(name: String) -> Sheet? {
        sheets.first { $0.name == name }
    }
    
    /// Removes a sheet by name
    public func removeSheet(name: String) {
        sheets.removeAll { $0.name == name }
    }
    
    /// Adds an image to the workbook
    public func addImage(_ image: ExcelImage) {
        images.append(image)
    }
    
    /// Gets all images in the workbook
    public func getImages() -> [ExcelImage] {
        images
    }
    
    /// Saves the workbook to a file asynchronously
    public func save(to url: URL) async throws {
        try await XLKitUtils.generateXLSX(workbook: self, to: url)
    }
    
    /// Saves the workbook to a file synchronously
    public func saveSync(to url: URL) throws {
        try XLKitUtils.generateXLSXSync(workbook: self, to: url)
    }
}

/// Represents a worksheet in an Excel workbook
public final class Sheet: Equatable {
    public let name: String
    public let id: Int
    private var cells: [String: CellValue] = [:]
    private var mergedRanges: [CellRange] = []
    private var columnWidths: [Int: Double] = [:]
    private var rowHeights: [Int: Double] = [:]
    private var images: [String: ExcelImage] = [:] // coordinate -> image
    
    public init(name: String, id: Int) {
        self.name = name
        self.id = id
    }
    
    /// Sets a cell value
    @discardableResult
    public func setCell(_ coordinate: String, value: CellValue) -> Self {
        cells[coordinate.uppercased()] = value
        return self
    }
    
    /// Gets a cell value
    public func getCell(_ coordinate: String) -> CellValue? {
        cells[coordinate.uppercased()]
    }
    
    /// Sets a cell value by row and column
    @discardableResult
    public func setCell(row: Int, column: Int, value: CellValue) -> Self {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        setCell(coordinate, value: value)
        return self
    }
    
    /// Gets a cell value by row and column
    public func getCell(row: Int, column: Int) -> CellValue? {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        return getCell(coordinate)
    }
    
    /// Sets a range of cells with the same value
    @discardableResult
    public func setRange(_ range: String, value: CellValue) -> Self {
        guard let cellRange = CellRange(excelRange: range) else { return self }
        for coordinate in cellRange.coordinates {
            setCell(coordinate.excelAddress, value: value)
        }
        return self
    }
    
    /// Merges a range of cells
    @discardableResult
    public func mergeCells(_ range: String) -> Self {
        guard let cellRange = CellRange(excelRange: range) else { return self }
        mergedRanges.append(cellRange)
        return self
    }
    
    /// Gets all cell coordinates that have values
    public func getUsedCells() -> [String] {
        Array(cells.keys).sorted()
    }
    
    /// Gets all merged ranges
    public func getMergedRanges() -> [CellRange] {
        mergedRanges
    }
    
    /// Sets column width in pixels
    @discardableResult
    public func setColumnWidth(_ column: Int, width: Double) -> Self {
        columnWidths[column] = width
        return self
    }
    
    /// Gets column width in pixels
    public func getColumnWidth(_ column: Int) -> Double? {
        columnWidths[column]
    }
    
    /// Gets all column widths
    public func getColumnWidths() -> [Int: Double] {
        columnWidths
    }
    
    /// Sets row height in pixels
    @discardableResult
    public func setRowHeight(_ row: Int, height: Double) -> Self {
        rowHeights[row] = height
        return self
    }
    
    /// Gets row height in pixels
    public func getRowHeight(_ row: Int) -> Double? {
        rowHeights[row]
    }
    
    /// Gets all row heights
    public func getRowHeights() -> [Int: Double] {
        rowHeights
    }
    
    /// Auto-sizes a column to fit an image
    @discardableResult
    public func autoSizeColumn(_ column: Int, forImageAt coordinate: String) -> Self {
        guard let image = images[coordinate] else { return self }
        let width = image.displaySize?.width ?? image.originalSize.width
        setColumnWidth(column, width: width)
        return self
    }
    
    /// Adds an image to a cell
    @discardableResult
    public func addImage(_ image: ExcelImage, at coordinate: String) -> Self {
        images[coordinate.uppercased()] = image
        return self
    }
    
    /// Adds an image from Data
    @discardableResult
    public func addImage(_ data: Data, at coordinate: String, format: ImageFormat, displaySize: CGSize? = nil) -> Bool {
        guard let image = ExcelImage.from(data: data, format: format, displaySize: displaySize) else { return false }
        addImage(image, at: coordinate)
        return true
    }
    
    /// Adds an image from file URL
    @discardableResult
    public func addImage(from url: URL, at coordinate: String, displaySize: CGSize? = nil) throws -> Bool {
        guard let image = try ExcelImage.from(url: url, displaySize: displaySize) else { return false }
        addImage(image, at: coordinate)
        return true
    }
    
    /// Gets all images in the sheet
    public func getImages() -> [String: ExcelImage] {
        images
    }
    
    /// Clears all data from the sheet
    public func clear() {
        cells.removeAll()
        mergedRanges.removeAll()
        columnWidths.removeAll()
        rowHeights.removeAll()
        images.removeAll()
    }
    
    public static func == (lhs: Sheet, rhs: Sheet) -> Bool {
        lhs === rhs // Reference equality
    }
}

/// Represents a cell coordinate (row, column)
public struct CellCoordinate: Hashable {
    public let row: Int
    public let column: Int
    
    public init(row: Int, column: Int) {
        self.row = row
        self.column = column
    }
    
    /// Excel-style address (e.g., "A1", "B2")
    public var excelAddress: String {
        return XLKitUtils.columnLetter(from: column) + String(row)
    }
    
    /// Creates a coordinate from Excel address
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
        let column = XLKitUtils.columnNumber(from: columnString)
        
        self.row = row
        self.column = column
    }
}

/// Represents a range of cells
public struct CellRange: Hashable {
    public let start: CellCoordinate
    public let end: CellCoordinate
    
    public init(start: CellCoordinate, end: CellCoordinate) {
        self.start = start
        self.end = end
    }
    
    /// Creates a range from Excel range notation (e.g., "A1:B5")
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
    
    /// All coordinates in this range
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
            return XLKitUtils.formatDate(value)
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

/// Error types for XLKit operations
public enum XLKitError: Error, LocalizedError {
    case invalidCoordinate(String)
    case invalidRange(String)
    case fileWriteError(String)
    case zipCreationError(String)
    case xmlGenerationError(String)
    
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
        }
    }
}
