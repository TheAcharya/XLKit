import Foundation

public struct XCS{
    private static var table: [UInt32] = {
        (0...255).map { i -> UInt32 in
            (0..<8).reduce(UInt32(i), { c, _ in
                (c % 2 == 0) ? (c >> 1) : (0xEDB88320 ^ (c >> 1))
            })
        }
    }()
    
    public static func checksum(data: Data) -> UInt32 {
        var byteData = [UInt8](repeating:0, count: data.count)
        data.copyBytes(to: &byteData, count: data.count)
        return checksum(bytes: byteData)
    }
    
    /// check sum CRC32
    public static func checksum(bytes: [UInt8]) -> UInt32 {
        return ~(bytes.reduce(~UInt32(0), { crc, byte in
            (crc >> 8) ^ table[(Int(crc) ^ Int(byte)) & 0xFF]
        }))
    }
}
