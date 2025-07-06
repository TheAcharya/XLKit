import Foundation

/// Universal Excel library for creating and manipulating .xlsx files
public struct XLKit {
    
    /// Creates a new workbook
    public static func createWorkbook() -> Workbook {
        return Workbook()
    }
    
    /// Saves a workbook to a file asynchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) async throws {
        try await workbook.save(to: url)
    }
    
    /// Saves a workbook to a file synchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) throws {
        try workbook.saveSync(to: url)
    }
    
    // MARK: - CSV/TSV Convenience Methods
    
    /// Creates a workbook from CSV data
    public static func createWorkbookFromCSV(csvData: String, sheetName: String = "Sheet1", separator: String = ",", hasHeader: Bool = false) -> Workbook {
        return CSVUtils.createWorkbookFromCSV(csvData: csvData, sheetName: sheetName, separator: separator, hasHeader: hasHeader)
    }
    
    /// Creates a workbook from TSV data
    public static func createWorkbookFromTSV(tsvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) -> Workbook {
        return CSVUtils.createWorkbookFromTSV(tsvData: tsvData, sheetName: sheetName, hasHeader: hasHeader)
    }
    
    /// Exports a sheet to CSV format
    public static func exportSheetToCSV(sheet: Sheet, separator: String = ",") -> String {
        return CSVUtils.exportToCSV(sheet: sheet, separator: separator)
    }
    
    /// Exports a sheet to TSV format
    public static func exportSheetToTSV(sheet: Sheet) -> String {
        return CSVUtils.exportToTSV(sheet: sheet)
    }
    
    /// Imports CSV data into a sheet
    public static func importCSVIntoSheet(sheet: Sheet, csvData: String, separator: String = ",", hasHeader: Bool = false) {
        CSVUtils.importFromCSV(sheet: sheet, csvData: csvData, separator: separator, hasHeader: hasHeader)
    }
    
    /// Imports TSV data into a sheet
    public static func importTSVIntoSheet(sheet: Sheet, tsvData: String, hasHeader: Bool = false) {
        CSVUtils.importFromTSV(sheet: sheet, tsvData: tsvData, hasHeader: hasHeader)
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
    private var cellFormats: [String: CellFormat] = [:] // coordinate -> format
    
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
    
    /// Gets a cell with formatting
    public func getCellWithFormat(_ coordinate: String) -> Cell? {
        guard let value = cells[coordinate.uppercased()] else { return nil }
        let format = cellFormats[coordinate.uppercased()]
        return Cell(value, format: format)
    }
    
    /// Gets cell formatting
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
    
    /// Sets a cell with string value and formatting
    @discardableResult
    public func setCell(_ coordinate: String, string: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.string(string, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with number value and formatting
    @discardableResult
    public func setCell(_ coordinate: String, number: Double, format: CellFormat? = nil) -> Self {
        let cell = Cell.number(number, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with integer value and formatting
    @discardableResult
    public func setCell(_ coordinate: String, integer: Int, format: CellFormat? = nil) -> Self {
        let cell = Cell.integer(integer, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with boolean value and formatting
    @discardableResult
    public func setCell(_ coordinate: String, boolean: Bool, format: CellFormat? = nil) -> Self {
        let cell = Cell.boolean(boolean, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with date value and formatting
    @discardableResult
    public func setCell(_ coordinate: String, date: Date, format: CellFormat? = nil) -> Self {
        let cell = Cell.date(date, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a cell with formula and formatting
    @discardableResult
    public func setCell(_ coordinate: String, formula: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.formula(formula, format: format)
        return setCell(coordinate, cell: cell)
    }
    
    /// Sets a range of cells with the same value and formatting
    @discardableResult
    public func setRange(_ range: String, cell: Cell) -> Self {
        guard let cellRange = CellRange(excelRange: range) else { return self }
        for coordinate in cellRange.coordinates {
            setCell(coordinate.excelAddress, cell: cell)
        }
        return self
    }
    
    /// Sets a range of cells with string value and formatting
    @discardableResult
    public func setRange(_ range: String, string: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.string(string, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with number value and formatting
    @discardableResult
    public func setRange(_ range: String, number: Double, format: CellFormat? = nil) -> Self {
        let cell = Cell.number(number, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with integer value and formatting
    @discardableResult
    public func setRange(_ range: String, integer: Int, format: CellFormat? = nil) -> Self {
        let cell = Cell.integer(integer, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with boolean value and formatting
    @discardableResult
    public func setRange(_ range: String, boolean: Bool, format: CellFormat? = nil) -> Self {
        let cell = Cell.boolean(boolean, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with date value and formatting
    @discardableResult
    public func setRange(_ range: String, date: Date, format: CellFormat? = nil) -> Self {
        let cell = Cell.date(date, format: format)
        return setRange(range, cell: cell)
    }
    
    /// Sets a range of cells with formula and formatting
    @discardableResult
    public func setRange(_ range: String, formula: String, format: CellFormat? = nil) -> Self {
        let cell = Cell.formula(formula, format: format)
        return setRange(range, cell: cell)
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
        cellFormats.removeAll()
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
