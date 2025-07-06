// swift-tools-version: 6.0
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
        )
    ],
    targets: [
        .target(
            name: "XLKit"
        ),
        .testTarget(
            name: "XLKitTests",
            dependencies: ["XLKit"]
        )
    ]
)
