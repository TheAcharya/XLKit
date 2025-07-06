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
}

/// Represents an Excel workbook containing multiple sheets
public class Workbook {
    private var sheets: [Sheet] = []
    private let nextSheetId: Int
    
    public init() {
        self.nextSheetId = 1
    }
    
    /// Adds a new sheet to the workbook
    public func addSheet(name: String) -> Sheet {
        let sheet = Sheet(name: name, id: nextSheetId)
        sheets.append(sheet)
        return sheet
    }
    
    /// Gets all sheets in the workbook
    public func getSheets() -> [Sheet] {
        return sheets
    }
    
    /// Gets a sheet by name
    public func getSheet(name: String) -> Sheet? {
        return sheets.first { $0.name == name }
    }
    
    /// Removes a sheet by name
    public func removeSheet(name: String) {
        sheets.removeAll { $0.name == name }
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
public class Sheet {
    public let name: String
    public let id: Int
    private var cells: [String: CellValue] = [:]
    private var mergedRanges: [CellRange] = []
    
    public init(name: String, id: Int) {
        self.name = name
        self.id = id
    }
    
    /// Sets a cell value
    public func setCell(_ coordinate: String, value: CellValue) {
        cells[coordinate.uppercased()] = value
    }
    
    /// Gets a cell value
    public func getCell(_ coordinate: String) -> CellValue? {
        return cells[coordinate.uppercased()]
    }
    
    /// Sets a cell value by row and column
    public func setCell(row: Int, column: Int, value: CellValue) {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        setCell(coordinate, value: value)
    }
    
    /// Gets a cell value by row and column
    public func getCell(row: Int, column: Int) -> CellValue? {
        let coordinate = CellCoordinate(row: row, column: column).excelAddress
        return getCell(coordinate)
    }
    
    /// Sets a range of cells with the same value
    public func setRange(_ range: String, value: CellValue) {
        guard let cellRange = CellRange(excelRange: range) else { return }
        for coordinate in cellRange.coordinates {
            setCell(coordinate.excelAddress, value: value)
        }
    }
    
    /// Merges a range of cells
    public func mergeCells(_ range: String) {
        guard let cellRange = CellRange(excelRange: range) else { return }
        mergedRanges.append(cellRange)
    }
    
    /// Gets all cell coordinates that have values
    public func getUsedCells() -> [String] {
        return Array(cells.keys).sorted()
    }
    
    /// Gets all merged ranges
    public func getMergedRanges() -> [CellRange] {
        return mergedRanges
    }
    
    /// Clears all data from the sheet
    public func clear() {
        cells.removeAll()
        mergedRanges.removeAll()
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
