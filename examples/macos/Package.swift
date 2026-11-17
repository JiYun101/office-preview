// swift-tools-version: 5.9
import PackageDescription
import Foundation

let infoPlist = URL(fileURLWithPath: #filePath)
    .deletingLastPathComponent().appendingPathComponent("Info.plist").path

let package = Package(
    name: "OfficePreviewMac",
    platforms: [.macOS(.v14)],
    products: [.executable(name: "OfficePreviewMac", targets: ["OfficePreviewMac"])],
    targets: [
        .executableTarget(
            name: "OfficePreviewMac",
            resources: [.copy("WebAssets")],
            linkerSettings: [
                .unsafeFlags(["-Xlinker", "-sectcreate", "-Xlinker", "__TEXT", "-Xlinker", "__info_plist", "-Xlinker", infoPlist])
            ]
        )
    ]
)
