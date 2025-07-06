// swift-tools-version:5.5
// The swift-tools-version declares the minimum version of Swift required to build this package.

import PackageDescription

let package = Package(
    name: "XLKit",
    platforms: [
        .iOS(.v10),
        .macOS(.v11),
    ],
    products: [
        .library(
            name: "XLKit",
            targets: ["XLKit"]),
    ],
    dependencies: [
        .package(url: "https://github.com/ZipArchive/ZipArchive.git", "2.3.0"..<"2.5.0")
    ],
    targets: [
        .target(
            name: "XLKit",
            dependencies: ["ZipArchive"]),
        .testTarget(
            name: "XLKitTests",
            dependencies: ["XLKit"]),
    ]
)
