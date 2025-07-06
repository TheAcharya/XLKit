//
//  ImageUtils.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import CoreGraphics
@preconcurrency import XLKitCore

// MARK: - Image Utilities for XLKit

/// Image utility functions for XLKit
public struct ImageUtils {
    
    // MARK: - Image Format Detection
    
    /// Detects image format from data
    public static func detectImageFormat(from data: Data) -> ImageFormat? {
        guard data.count >= 4 else { return nil }
        let bytes = [UInt8](data)
        
        // Check GIF signature
        if bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 {
            return .gif
        }
        
        // Check PNG signature
        if bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 {
            return .png
        }
        
        // Check JPEG signature
        if bytes[0] == 0xFF && bytes[1] == 0xD8 {
            return .jpeg
        }
        
        // Check BMP signature
        if bytes[0] == 0x42 && bytes[1] == 0x4D {
            return .bmp
        }
        
        // Check TIFF signature (both little and big endian)
        if (bytes[0] == 0x49 && bytes[1] == 0x49) || (bytes[0] == 0x4D && bytes[1] == 0x4D) {
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
        
        // GIF header: width (2 bytes) + height (2 bytes)
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
        guard data.count >= 2 else { return nil }
        let bytes = [UInt8](data)
        
        // JPEG starts with 0xFF 0xD8
        guard bytes[0] == 0xFF && bytes[1] == 0xD8 else { return nil }
        
        var i = 2
        while i < data.count - 1 {
            // Look for SOF markers (0xFF 0xC0-0xCF, except 0xC4, 0xC8, 0xCC)
            if bytes[i] == 0xFF && bytes[i + 1] >= 0xC0 && bytes[i + 1] <= 0xCF {
                if bytes[i + 1] != 0xC4 && bytes[i + 1] != 0xC8 && bytes[i + 1] != 0xCC {
                    if i + 9 < data.count {
                        let height = Int(bytes[i + 5]) << 8 | Int(bytes[i + 6])
                        let width = Int(bytes[i + 7]) << 8 | Int(bytes[i + 8])
                        return CGSize(width: Double(width), height: Double(height))
                    }
                }
            }
            i += 1
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