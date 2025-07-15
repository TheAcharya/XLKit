//
//  SecurityManager.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation

@MainActor
public struct SecurityManager {
    
    // MARK: - Configuration
    
    /// Flag to enable/disable checksum file creation
    public static var enableChecksumStorage = false
    
    /// Flag to enable/disable file path security restrictions
    /// When disabled (default), files can be saved to any location
    /// When enabled, files are restricted to temporary directory and home directory
    public static var enableFilePathRestrictions = false
    
    // MARK: - Rate Limiting
    
    /// Rate limiter for file operations (100 per minute)
    private static var operationLimiter = RateLimiter(maxOperations: 100, timeWindow: 60)
    
    /// Checks if operation is allowed under rate limiting
    public static func checkRateLimit() throws {
        guard operationLimiter.allowOperation() else {
            logSecurityOperation("rate_limit_exceeded", details: [
                "timestamp": Date().timeIntervalSince1970,
                "operation": "file_operation"
            ])
            throw XLKitError.rateLimitExceeded("Rate limit exceeded: maximum 100 operations per minute")
        }
    }
    
    // MARK: - Security Logging
    
    /// Security log entry
    public struct SecurityLogEntry {
        public let timestamp: Date
        public let operation: String
        public let details: [String: Any]
        public let userAgent: String
        
        public init(timestamp: Date, operation: String, details: [String: Any], userAgent: String) {
            self.timestamp = timestamp
            self.operation = operation
            self.details = details
            self.userAgent = userAgent
        }
    }
    
    /// Logs security-relevant operations
    public static func logSecurityOperation(_ operation: String, details: [String: Any]) {
        let logEntry = SecurityLogEntry(
            timestamp: Date(),
            operation: operation,
            details: details,
            userAgent: "XLKit/1.0.1"
        )
        
        // Log to console for development (in production, this would go to a secure audit trail)
        print("[SECURITY] \(logEntry.timestamp): \(operation) - \(details)")
        
        // Store in security log file
        storeSecurityLog(logEntry)
    }
    
    /// Stores security log entry
    private static func storeSecurityLog(_ entry: SecurityLogEntry) {
        let logDir = FileManager.default.temporaryDirectory.appendingPathComponent("xlkit_security_logs")
        try? FileManager.default.createDirectory(at: logDir, withIntermediateDirectories: true)
        
        let logFile = logDir.appendingPathComponent("security.log")
        let logLine = "\(entry.timestamp): \(entry.operation) - \(entry.details)\n"
        
        if let data = logLine.data(using: .utf8) {
            if FileManager.default.fileExists(atPath: logFile.path) {
                if let fileHandle = try? FileHandle(forWritingTo: logFile) {
                    fileHandle.seekToEndOfFile()
                    fileHandle.write(data)
                    fileHandle.closeFile()
                }
            } else {
                try? data.write(to: logFile)
            }
        }
    }
    
    // MARK: - File Quarantine
    
    /// Quarantines suspicious files
    public static func quarantineSuspiciousFile(_ fileURL: URL, reason: String) throws {
        let quarantineDir = FileManager.default.temporaryDirectory.appendingPathComponent("xlkit_quarantine")
        try FileManager.default.createDirectory(at: quarantineDir, withIntermediateDirectories: true)
        
        let quarantinedURL = quarantineDir.appendingPathComponent(fileURL.lastPathComponent)
        try FileManager.default.moveItem(at: fileURL, to: quarantinedURL)
        
        // Log quarantine action
        logSecurityOperation("file_quarantined", details: [
            "original_path": fileURL.path,
            "quarantine_path": quarantinedURL.path,
            "reason": reason,
            "timestamp": Date().timeIntervalSince1970
        ])
        
        throw XLKitError.suspiciousFileDetected("File quarantined: \(reason)")
    }
    
    /// Checks if file should be quarantined
    public static func shouldQuarantineFile(_ data: Data, format: ImageFormat) -> Bool {
        // Check for suspicious patterns
        let suspiciousPatterns = [
            "eval(", "exec(", "system(", "shell_exec(",
            "javascript:", "vbscript:", "data:text/html"
        ]
        
        let dataString = String(data: data, encoding: .utf8) ?? ""
        
        for pattern in suspiciousPatterns {
            if dataString.lowercased().contains(pattern.lowercased()) {
                return true
            }
        }
        
        // Check for unusually large files for their format
        let maxSizes: [ImageFormat: Int] = [
            .gif: 10 * 1024 * 1024,  // 10MB for GIF
            .png: 20 * 1024 * 1024,  // 20MB for PNG
            .jpeg: 15 * 1024 * 1024, // 15MB for JPEG
            .jpg: 15 * 1024 * 1024   // 15MB for JPG
        ]
        
        if let maxSize = maxSizes[format], data.count > maxSize {
            return true
        }
        
        return false
    }
    
    // MARK: - File Checksums
    
    /// Stores file checksum
    public static func storeFileChecksum(_ checksum: String, for fileURL: URL) {
        // Skip checksum storage if disabled
        guard enableChecksumStorage else {
            return
        }
        
        let metadataURL = fileURL.appendingPathExtension("checksum")
        try? checksum.write(to: metadataURL, atomically: true, encoding: .utf8)
        
        logSecurityOperation("checksum_stored", details: [
            "file_path": fileURL.path,
            "checksum": checksum,
            "timestamp": Date().timeIntervalSince1970
        ])
    }
    
    /// Verifies file checksum
    public static func verifyFileChecksum(_ fileURL: URL) -> Bool {
        // Skip checksum verification if disabled
        guard enableChecksumStorage else {
            return true // Assume valid if checksums are disabled
        }
        
        let metadataURL = fileURL.appendingPathExtension("checksum")
        
        guard let storedChecksum = try? String(contentsOf: metadataURL, encoding: .utf8) else {
            return false
        }
        
        guard let currentChecksum = try? CoreUtils.generateFileChecksum(fileURL) else {
            return false
        }
        
        let isValid = storedChecksum == currentChecksum
        
        if !isValid {
            logSecurityOperation("checksum_mismatch", details: [
                "file_path": fileURL.path,
                "stored_checksum": storedChecksum,
                "current_checksum": currentChecksum,
                "timestamp": Date().timeIntervalSince1970
            ])
        }
        
        return isValid
    }
}

// MARK: - Rate Limiter

/// Simple rate limiter for security
private struct RateLimiter {
    private let maxOperations: Int
    private let timeWindow: TimeInterval
    private var operations: [Date] = []
    private let queue = DispatchQueue(label: "xlkit.rate.limiter")
    
    init(maxOperations: Int, timeWindow: TimeInterval) {
        self.maxOperations = maxOperations
        self.timeWindow = timeWindow
    }
    
    mutating func allowOperation() -> Bool {
        return queue.sync {
            let now = Date()
            let cutoff = now.addingTimeInterval(-timeWindow)
            
            // Remove old operations
            operations = operations.filter { $0 > cutoff }
            
            // Check if we can allow this operation
            guard operations.count < maxOperations else {
                return false
            }
            
            // Add current operation
            operations.append(now)
            return true
        }
    }
} 