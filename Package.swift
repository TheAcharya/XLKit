// swift-tools-version: 6.0
// The swift-tools-version declares the minimum version of Swift required to build this package.

import PackageDescription

let package = Package(
    name: "XLKit",
    platforms: [
        .macOS(.v12),
        .iOS(.v15)
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
    dependencies: [
        .package(url: "https://github.com/CoreOffice/CoreXLSX.git", from: "0.14.2"),
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.20"),
        .package(url: "https://github.com/orchetect/swift-textfile.git", from: "0.5.1")
    ],
    targets: [
        .target(
            name: "XLKitCore"
        ),
        .target(
            name: "XLKitFormatters",
            dependencies: ["XLKitCore", .product(name: "TextFile", package: "swift-textfile")]
        ),
        .target(
            name: "XLKitImages",
            dependencies: ["XLKitCore"]
        ),
        .target(
            name: "XLKitXLSX",
            dependencies: ["XLKitCore", "XLKitFormatters", "XLKitImages", "ZIPFoundation"]
        ),
        .target(
            name: "XLKit",
            dependencies: ["XLKitCore", "XLKitFormatters", "XLKitImages", "XLKitXLSX"]
        ),
        .testTarget(
            name: "XLKitTests",
            dependencies: ["XLKit"],
            swiftSettings: [
                .unsafeFlags(["-Xfrontend", "-disable-round-trip-debug-types"])
            ]
        ),
        .executableTarget(
            name: "XLKitTestRunner",
            dependencies: ["XLKit", "CoreXLSX"],
            exclude: ["README.md", "Templates"],
            swiftSettings: [
                .unsafeFlags(["-Xfrontend", "-disable-round-trip-debug-types"]),
                .unsafeFlags(["-Xfrontend", "-disable-availability-checking"])
            ]
        )
    ]
)
