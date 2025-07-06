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
        let relsDir = tempDir.appendingPathComponent("_rels")
        let xlRelsDir = xlDir.appendingPathComponent("_rels")
        
        try FileManager.default.createDirectory(at: xlDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: worksheetsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: relsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: xlRelsDir, withIntermediateDirectories: true)
        
        // Generate content files
        try generateContentTypes(tempDir: tempDir, workbook: workbook)
        try generateRels(tempDir: tempDir, workbook: workbook)
        try generateWorkbookRels(xlRelsDir: xlRelsDir, workbook: workbook)
        try generateWorkbook(xlDir: xlDir, workbook: workbook)
        try generateWorksheets(worksheetsDir: worksheetsDir, workbook: workbook)
        
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
        }
    }
    
    private static func generateWorksheetXML(sheet: Sheet) -> String {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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
        </worksheet>
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