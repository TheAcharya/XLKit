import Foundation

/// Utility functions for XLKit
public struct XLKitUtils {
    
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
    
    // MARK: - Image Utilities
    
    /// Detects image format from data
    public static func detectImageFormat(from data: Data) -> ImageFormat? {
        let bytes = [UInt8](data.prefix(8))
        
        // GIF (needs 6 bytes)
        if bytes.count >= 6 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 {
            return .gif
        }
        // PNG (needs 8 bytes)
        if bytes.count >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 {
            return .png
        }
        // JPEG (needs 2 bytes)
        if bytes.count >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8 {
            return .jpeg
        }
        // BMP (needs 2 bytes)
        if bytes.count >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4D {
            return .bmp
        }
        // TIFF (needs 4 bytes)
        if bytes.count >= 4 && ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0x00) ||
                                (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00 && bytes[3] == 0x2A)) {
            return .tiff
        }
        return nil
    }
    
    /// Gets image size from data
    public static func getImageSize(from data: Data, format: ImageFormat) -> CGSize? {
        switch format {
        case .gif:
            return getGIFSize(from: data)
        case .png:
            return getPNGSize(from: data)
        case .jpeg, .jpg:
            return getJPEGSize(from: data)
        case .bmp:
            return getBMPSize(from: data)
        case .tiff:
            return getTIFFSize(from: data)
        }
    }
    
    /// Gets GIF image size
    private static func getGIFSize(from data: Data) -> CGSize? {
        guard data.count >= 10 else { return nil }
        let bytes = [UInt8](data)
        
        // GIF header: 6 bytes + width (2 bytes) + height (2 bytes)
        let width = Int(bytes[6]) | (Int(bytes[7]) << 8)
        let height = Int(bytes[8]) | (Int(bytes[9]) << 8)
        
        return CGSize(width: Double(width), height: Double(height))
    }
    
    /// Gets PNG image size
    private static func getPNGSize(from data: Data) -> CGSize? {
        guard data.count >= 24 else { return nil }
        let bytes = [UInt8](data)
        
        // PNG IHDR chunk: width (4 bytes) + height (4 bytes)
        let width = Int(bytes[16]) << 24 | Int(bytes[17]) << 16 | Int(bytes[18]) << 8 | Int(bytes[19])
        let height = Int(bytes[20]) << 24 | Int(bytes[21]) << 16 | Int(bytes[22]) << 8 | Int(bytes[23])
        
        return CGSize(width: Double(width), height: Double(height))
    }
    
    /// Gets JPEG image size
    private static func getJPEGSize(from data: Data) -> CGSize? {
        var offset = 2 // Skip SOI marker
        
        while offset < data.count - 1 {
            guard data[offset] == 0xFF else { return nil }
            
            let marker = data[offset + 1]
            if marker == 0xC0 || marker == 0xC1 || marker == 0xC2 {
                // SOF marker found
                guard offset + 9 < data.count else { return nil }
                
                let height = Int(data[offset + 5]) << 8 | Int(data[offset + 6])
                let width = Int(data[offset + 7]) << 8 | Int(data[offset + 8])
                
                return CGSize(width: Double(width), height: Double(height))
            }
            
            // Skip to next marker
            guard offset + 3 < data.count else { return nil }
            let length = Int(data[offset + 2]) << 8 | Int(data[offset + 3])
            offset += length + 2
        }
        
        return nil
    }
    
    /// Gets BMP image size
    private static func getBMPSize(from data: Data) -> CGSize? {
        guard data.count >= 26 else { return nil }
        let bytes = [UInt8](data)
        
        // BMP header: width (4 bytes) + height (4 bytes)
        let width = Int(bytes[18]) | (Int(bytes[19]) << 8) | (Int(bytes[20]) << 16) | (Int(bytes[21]) << 24)
        let height = Int(bytes[22]) | (Int(bytes[23]) << 8) | (Int(bytes[24]) << 16) | (Int(bytes[25]) << 24)
        
        return CGSize(width: Double(width), height: Double(height))
    }
    
    /// Gets TIFF image size
    private static func getTIFFSize(from data: Data) -> CGSize? {
        guard data.count >= 8 else { return nil }
        _ = [UInt8](data)
        
        // TIFF is complex, for simplicity we'll return a default size
        // In a production environment, you'd want a proper TIFF parser
        return CGSize(width: 100, height: 100)
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
    
    // MARK: - XLSX Generation
    
    /// Generates XLSX file asynchronously
    public static func generateXLSX(workbook: Workbook, to url: URL) async throws {
        // For now, just call the sync version directly since file I/O is not CPU-bound
        // and the workbook is not shared across threads
        try generateXLSXSync(workbook: workbook, to: url)
    }
    
    /// Generates XLSX file synchronously
    public static func generateXLSXSync(workbook: Workbook, to url: URL) throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("xlkit-\(UUID().uuidString)")
        
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer {
            try? FileManager.default.removeItem(at: tempDir)
        }
        
        // Create required directories
        let xlDir = tempDir.appendingPathComponent("xl")
        let worksheetsDir = xlDir.appendingPathComponent("worksheets")
        let mediaDir = xlDir.appendingPathComponent("media")
        let relsDir = tempDir.appendingPathComponent("_rels")
        let xlRelsDir = xlDir.appendingPathComponent("_rels")
        let worksheetRelsDir = worksheetsDir.appendingPathComponent("_rels")
        
        try FileManager.default.createDirectory(at: xlDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: worksheetsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: mediaDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: relsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: xlRelsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: worksheetRelsDir, withIntermediateDirectories: true)
        
        // Generate content files
        try generateContentTypes(tempDir: tempDir, workbook: workbook)
        try generateRels(tempDir: tempDir, workbook: workbook)
        try generateWorkbookRels(xlRelsDir: xlRelsDir, workbook: workbook)
        try generateWorkbook(xlDir: xlDir, workbook: workbook)
        try generateWorksheets(worksheetsDir: worksheetsDir, workbook: workbook)
        try generateMediaFiles(mediaDir: mediaDir, workbook: workbook)
        
        // Create ZIP archive
        try createZIPArchive(from: tempDir, to: url)
    }
    
    // MARK: - Private XLSX Generation Methods
    
    private static func generateContentTypes(tempDir: URL, workbook: Workbook) throws {
        var content = """
        <?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
        <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
            <Default Extension=\"xml\" ContentType=\"application/xml\"/>
            <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
            <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
        """
        
        for sheet in workbook.getSheets() {
            content += """
            
            <Override PartName=\"/xl/worksheets/sheet\(sheet.id).xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
            """
        }
        
        // Add image content types
        for image in workbook.getImages() {
            content += """
            
            <Override PartName=\"/xl/media/\(image.id).\(image.format.rawValue)\" ContentType=\"\(image.format.excelContentType)\"/>
            """
        }
        
        content += """
        
        </Types>
        """
        
        try content.write(to: tempDir.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)
    }
    
    private static func generateRels(tempDir: URL, workbook: Workbook) throws {
        let content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """
        
        try content.write(to: tempDir.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorkbookRels(xlRelsDir: URL, workbook: Workbook) throws {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
        
        for sheet in workbook.getSheets() {
            content += """
            
            <Relationship Id="rId\(sheet.id)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet\(sheet.id).xml"/>
            """
        }
        
        content += """
        
        </Relationships>
        """
        
        try content.write(to: xlRelsDir.appendingPathComponent("workbook.xml.rels"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorkbook(xlDir: URL, workbook: Workbook) throws {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets>
        """
        
        for sheet in workbook.getSheets() {
            content += """
            
                <sheet name="\(escapeXML(sheet.name))" sheetId="\(sheet.id)" r:id="rId\(sheet.id)"/>
            """
        }
        
        content += """
        
            </sheets>
        </workbook>
        """
        
        try content.write(to: xlDir.appendingPathComponent("workbook.xml"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorksheets(worksheetsDir: URL, workbook: Workbook) throws {
        for sheet in workbook.getSheets() {
            let content = generateWorksheetXML(sheet: sheet)
            try content.write(to: worksheetsDir.appendingPathComponent("sheet\(sheet.id).xml"), atomically: true, encoding: .utf8)
            
            // Generate worksheet relationships if there are images
            if !sheet.getImages().isEmpty {
                try generateWorksheetRels(worksheetRelsDir: worksheetsDir.appendingPathComponent("_rels"), sheet: sheet)
            }
        }
    }
    
    private static func generateWorksheetRels(worksheetRelsDir: URL, sheet: Sheet) throws {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
        
        var imageId = 1
        for (_, image) in sheet.getImages() {
            content += """
            
            <Relationship Id="rId\(imageId)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/\(image.id).\(image.format.rawValue)"/>
            """
            imageId += 1
        }
        
        content += """
        
        </Relationships>
        """
        
        try content.write(to: worksheetRelsDir.appendingPathComponent("sheet\(sheet.id).xml.rels"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorksheetXML(sheet: Sheet) -> String {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        """
        
        // Add column widths if any
        if !sheet.getColumnWidths().isEmpty {
            content += generateColumnWidthsXML(sheet: sheet)
        }
        
        // Add row heights if any
        if !sheet.getRowHeights().isEmpty {
            content += generateRowHeightsXML(sheet: sheet)
        }
        
        content += """
        
            <sheetData>
        """
        
        // Group cells by row
        var rows: [Int: [String: CellValue]] = [:]
        for (coordinate, value) in sheet.getUsedCells().compactMap({ ($0, sheet.getCell($0)) }) {
            guard let cellCoord = CellCoordinate(excelAddress: coordinate) else { continue }
            if rows[cellCoord.row] == nil {
                rows[cellCoord.row] = [:]
            }
            rows[cellCoord.row]?[coordinate] = value
        }
        
        // Generate row XML
        for rowNum in rows.keys.sorted() {
            content += """
            
                <row r="\(rowNum)">
            """
            
            for coordinate in rows[rowNum]?.keys.sorted() ?? [] {
                guard let value = rows[rowNum]?[coordinate] else { continue }
                content += generateCellXML(coordinate: coordinate, value: value)
            }
            
            content += """
            
                </row>
            """
        }
        
        content += """
        
            </sheetData>
        """
        
        // Add images if any
        if !sheet.getImages().isEmpty {
            content += generateDrawingXML(sheet: sheet)
        }
        
        content += """
        
        </worksheet>
        """
        
        return content
    }
    
    private static func generateColumnWidthsXML(sheet: Sheet) -> String {
        var content = """
        
            <cols>
        """
        
        for (column, width) in sheet.getColumnWidths().sorted(by: { $0.key < $1.key }) {
            content += """
            
                <col min="\(column)" max="\(column)" width="\(width)" customWidth="1"/>
            """
        }
        
        content += """
        
            </cols>
        """
        
        return content
    }
    
    private static func generateRowHeightsXML(sheet: Sheet) -> String {
        var content = """
        
            <sheetFormatPr defaultRowHeight="15"/>
        """
        
        for (row, height) in sheet.getRowHeights().sorted(by: { $0.key < $1.key }) {
            content += """
            
            <row r="\(row)" ht="\(height)" customHeight="1"/>
            """
        }
        
        return content
    }
    
    private static func generateDrawingXML(sheet: Sheet) -> String {
        let content = """
        
            <drawing r:id="rId1"/>
        """
        
        return content
    }
    
    private static func generateCellXML(coordinate: String, value: CellValue) -> String {
        switch value {
        case .string(let stringValue):
            return """
            
                    <c r="\(coordinate)" t="s">
                        <v>\(escapeXML(stringValue))</v>
                    </c>
            """
        case .number(let numberValue):
            return """
            
                    <c r="\(coordinate)" t="n">
                        <v>\(numberValue)</v>
                    </c>
            """
        case .integer(let intValue):
            return """
            
                    <c r="\(coordinate)" t="n">
                        <v>\(intValue)</v>
                    </c>
            """
        case .boolean(let boolValue):
            return """
            
                    <c r="\(coordinate)" t="b">
                        <v>\(boolValue ? 1 : 0)</v>
                    </c>
            """
        case .date(let dateValue):
            let excelNumber = excelNumberFromDate(dateValue)
            return """
            
                    <c r="\(coordinate)" t="n">
                        <v>\(excelNumber)</v>
                    </c>
            """
        case .formula(let formulaValue):
            return """
            
                    <c r="\(coordinate)" t="str">
                        <f>\(escapeXML(formulaValue))</f>
                    </c>
            """
        case .empty:
            return """
            
                    <c r="\(coordinate)"/>
            """
        }
    }
    
    private static func generateMediaFiles(mediaDir: URL, workbook: Workbook) throws {
        for image in workbook.getImages() {
            let imageURL = mediaDir.appendingPathComponent("\(image.id).\(image.format.rawValue)")
            try image.data.write(to: imageURL)
        }
    }
    
    // MARK: - ZIP Archive Creation
    
    private static func createZIPArchive(from sourceDir: URL, to destinationURL: URL) throws {
        // Remove existing file if it exists
        if FileManager.default.fileExists(atPath: destinationURL.path) {
            try FileManager.default.removeItem(at: destinationURL)
        }
        
        // Use system zip command for simplicity
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/zip")
        process.arguments = ["-r", destinationURL.path, "."]
        process.currentDirectoryURL = sourceDir
        
        try process.run()
        process.waitUntilExit()
        
        if process.terminationStatus != 0 {
            throw XLKitError.zipCreationError("ZIP creation failed with status \(process.terminationStatus)")
        }
    }
}

// MARK: - CSV/TSV Import/Export

/// CSV/TSV import/export utilities
public struct CSVUtils {
    
    /// Exports a sheet to CSV format
    public static func exportToCSV(sheet: Sheet, separator: String = ",") -> String {
        var csv = ""
        
        // Get all used cells and determine the range
        let usedCells = sheet.getUsedCells()
        guard !usedCells.isEmpty else { return csv }
        
        // Parse coordinates to find max row and column
        var maxRow = 0
        var maxColumn = 0
        
        for cellAddress in usedCells {
            if let coord = CellCoordinate(excelAddress: cellAddress) {
                maxRow = max(maxRow, coord.row)
                maxColumn = max(maxColumn, coord.column)
            }
        }
        
        // Generate CSV content
        for row in 1...maxRow {
            var rowData: [String] = []
            
            for column in 1...maxColumn {
                let coord = CellCoordinate(row: row, column: column)
                let cellAddress = coord.excelAddress
                
                if let cellValue = sheet.getCell(cellAddress) {
                    let stringValue = formatCellValueForCSV(cellValue, separator: separator)
                    rowData.append(stringValue)
                } else {
                    rowData.append("")
                }
            }
            
            csv += rowData.joined(separator: separator) + "\n"
        }
        
        return csv
    }
    
    /// Exports a sheet to TSV format
    public static func exportToTSV(sheet: Sheet) -> String {
        return exportToCSV(sheet: sheet, separator: "\t")
    }
    
    /// Imports CSV data into a sheet
    public static func importFromCSV(sheet: Sheet, csvData: String, separator: String = ",", hasHeader: Bool = false) {
        let lines = csvData.components(separatedBy: .newlines)
            .filter { !$0.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty }
        
        var startRow = 1
        if hasHeader {
            startRow = 2 // Skip header row
        }
        
        for (lineIndex, line) in lines.enumerated() {
            let row = startRow + lineIndex
            let columns = parseCSVLine(line, separator: separator)
            
            for (columnIndex, value) in columns.enumerated() {
                let column = columnIndex + 1
                let coord = CellCoordinate(row: row, column: column)
                let cellAddress = coord.excelAddress
                
                let cellValue = parseCSVValue(value)
                sheet.setCell(cellAddress, value: cellValue)
            }
        }
    }
    
    /// Imports TSV data into a sheet
    public static func importFromTSV(sheet: Sheet, tsvData: String, hasHeader: Bool = false) {
        importFromCSV(sheet: sheet, csvData: tsvData, separator: "\t", hasHeader: hasHeader)
    }
    
    /// Creates a workbook from CSV data
    public static func createWorkbookFromCSV(csvData: String, sheetName: String = "Sheet1", separator: String = ",", hasHeader: Bool = false) -> Workbook {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: sheetName)
        
        importFromCSV(sheet: sheet, csvData: csvData, separator: separator, hasHeader: hasHeader)
        
        return workbook
    }
    
    /// Creates a workbook from TSV data
    public static func createWorkbookFromTSV(tsvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) -> Workbook {
        return createWorkbookFromCSV(csvData: tsvData, sheetName: sheetName, separator: "\t", hasHeader: hasHeader)
    }
    
    // MARK: - Private Helper Methods
    
    private static func formatCellValueForCSV(_ cellValue: CellValue, separator: String) -> String {
        let stringValue = cellValue.stringValue
        
        // Escape quotes and wrap in quotes if needed
        if stringValue.contains(separator) || stringValue.contains("\"") || stringValue.contains("\n") {
            let escaped = stringValue.replacingOccurrences(of: "\"", with: "\"\"")
            return "\"\(escaped)\""
        }
        
        return stringValue
    }
    
    private static func parseCSVLine(_ line: String, separator: String) -> [String] {
        var result: [String] = []
        var current = ""
        var inQuotes = false
        
        for char in line {
            if char == "\"" {
                inQuotes.toggle()
            } else if char == Character(separator) && !inQuotes {
                result.append(current.trimmingCharacters(in: .whitespaces))
                current = ""
            } else {
                current.append(char)
            }
        }
        
        result.append(current.trimmingCharacters(in: .whitespaces))
        return result
    }
    
    private static func parseCSVValue(_ value: String) -> CellValue {
        let trimmed = value.trimmingCharacters(in: .whitespaces)
        
        if trimmed.isEmpty {
            return .empty
        }
        
        // Try to parse as number
        if let doubleValue = Double(trimmed) {
            // Check if it's an integer
            if doubleValue.truncatingRemainder(dividingBy: 1) == 0 {
                return .integer(Int(doubleValue))
            }
            return .number(doubleValue)
        }
        
        // Try to parse as boolean
        let lowercased = trimmed.lowercased()
        if lowercased == "true" || lowercased == "false" {
            return .boolean(lowercased == "true")
        }
        
        // Try to parse as date (basic ISO format)
        if let date = parseDate(trimmed) {
            return .date(date)
        }
        
        // Default to string
        return .string(trimmed)
    }
    
    private static func parseDate(_ value: String) -> Date? {
        let formatters = [
            "yyyy-MM-dd",
            "MM/dd/yyyy",
            "dd/MM/yyyy",
            "yyyy-MM-dd HH:mm:ss",
            "MM/dd/yyyy HH:mm:ss"
        ]
        
        for format in formatters {
            let formatter = DateFormatter()
            formatter.dateFormat = format
            if let date = formatter.date(from: value) {
                return date
            }
        }
        
        return nil
    }
} 