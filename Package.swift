// swift-tools-version: 5.9
// The swift-tools-version declares the minimum version of Swift required to build this package.

import PackageDescription

let package = Package(
    name: "XLKit",
    platforms: [
        .macOS(.v13)
    ],
    products: [
        .library(
            name: "XLKit",
            targets: ["XLKit"]
        ),
        .library(
            name: "XLKitCore",
            targets: ["XLKitCore"]
        ),
        .library(
            name: "XLKitFormatters",
            targets: ["XLKitFormatters"]
        ),
        .library(
            name: "XLKitImages",
            targets: ["XLKitImages"]
        ),
        .library(
            name: "XLKitXLSX",
            targets: ["XLKitXLSX"]
        ),
    ],
    targets: [
        .target(
            name: "XLKitCore"
        ),
        .target(
            name: "XLKitFormatters",
            dependencies: ["XLKitCore"]
        ),
        .target(
            name: "XLKitImages",
            dependencies: ["XLKitCore"]
        ),
        .target(
            name: "XLKitXLSX",
            dependencies: ["XLKitCore", "XLKitFormatters", "XLKitImages"]
        ),
        .target(
            name: "XLKit",
            dependencies: ["XLKitCore", "XLKitFormatters", "XLKitImages", "XLKitXLSX"]
        ),
        .testTarget(
            name: "XLKitTests",
            dependencies: ["XLKit"]
        ),
        .executableTarget(
            name: "XLKitTestRunner",
            dependencies: ["XLKit"]
        )
    ]
)
