import Foundation
import CoreGraphics
import XLKitCore

// MARK: - Image Utilities for XLKit

/// Image utility functions for XLKit
public struct ImageUtils {
    
    // MARK: - Image Format Detection
    
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
    
    /// Creates an ExcelImage from Data
    public static func createExcelImage(from data: Data, format: ImageFormat, displaySize: CGSize? = nil) -> ExcelImage? {
        guard let size = getImageSize(from: data, format: format) else { return nil }
        let id = "image_\(UUID().uuidString)"
        return ExcelImage(id: id, data: data, format: format, originalSize: size, displaySize: displaySize)
    }
    
    /// Creates an ExcelImage from a file URL
    public static func createExcelImage(from url: URL, displaySize: CGSize? = nil) throws -> ExcelImage? {
        let data = try Data(contentsOf: url)
        guard let format = detectImageFormat(from: data) else { return nil }
        return createExcelImage(from: data, format: format, displaySize: displaySize)
    }
}

public extension Sheet {
    /// Adds an image from Data
    @discardableResult
    func addImage(_ data: Data, at coordinate: String, format: ImageFormat, displaySize: CGSize? = nil) -> Bool {
        guard let image = ImageUtils.createExcelImage(from: data, format: format, displaySize: displaySize) else { return false }
        addImage(image, at: coordinate)
        return true
    }
    /// Adds an image from file URL
    @discardableResult
    func addImage(from url: URL, at coordinate: String, displaySize: CGSize? = nil) throws -> Bool {
        guard let image = try ImageUtils.createExcelImage(from: url, displaySize: displaySize) else { return false }
        addImage(image, at: coordinate)
        return true
    }
} 