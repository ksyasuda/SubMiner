#!/usr/bin/env swift
//
//  get-mpv-window-macos.swift
//  SubMiner - Get mpv window geometry on macOS
//
//  This script uses Core Graphics APIs to find mpv windows system-wide.
//  It works with both bundled and unbundled mpv installations.
//
//  Usage: swift get-mpv-window-macos.swift
//  Output: "x,y,width,height" or "not-found"
//

import Cocoa
import Foundation

private struct WindowGeometry {
    let x: Int
    let y: Int
    let width: Int
    let height: Int
}

private struct WindowState {
    let geometry: WindowGeometry
    let focused: Bool
}

private enum WindowLookupResult {
    case visible(WindowState)
    case minimized
}

private let targetMpvSocketPath: String? = {
    guard CommandLine.arguments.count > 1 else {
        return nil
    }

    let value = CommandLine.arguments[1].trimmingCharacters(in: .whitespacesAndNewlines)
    return value.isEmpty ? nil : value
}()

private func windowCommandLineForPid(_ pid: pid_t) -> String? {
    let process = Process()
    process.executableURL = URL(fileURLWithPath: "/bin/ps")
    process.arguments = ["-p", "\(pid)", "-o", "args="]

    let pipe = Pipe()
    process.standardOutput = pipe
    process.standardError = pipe

    do {
        try process.run()
        process.waitUntilExit()
    } catch {
        return nil
    }

    if process.terminationStatus != 0 {
        return nil
    }

    let data = pipe.fileHandleForReading.readDataToEndOfFile()
    guard let commandLine = String(data: data, encoding: .utf8) else {
        return nil
    }

    return commandLine
        .trimmingCharacters(in: .whitespacesAndNewlines)
        .replacingOccurrences(of: "\n", with: " ")
}

private func windowHasTargetSocket(_ pid: pid_t) -> Bool {
    guard let socketPath = targetMpvSocketPath else {
        return true
    }

    guard let commandLine = windowCommandLineForPid(pid) else {
        return false
    }

    return commandLine.contains("--input-ipc-server=\(socketPath)") ||
        commandLine.contains("--input-ipc-server \(socketPath)")
}

private func geometryFromRect(x: CGFloat, y: CGFloat, width: CGFloat, height: CGFloat) -> WindowGeometry {
    let minX = Int(floor(x))
    let minY = Int(floor(y))
    let maxX = Int(ceil(x + width))
    let maxY = Int(ceil(y + height))
    return WindowGeometry(
        x: minX,
        y: minY,
        width: max(0, maxX - minX),
        height: max(0, maxY - minY)
    )
}

private func normalizedMpvName(_ name: String) -> Bool {
    let normalized = name.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
    return normalized == "mpv"
}

private func geometryFromAXWindow(_ axWindow: AXUIElement) -> WindowGeometry? {
    var positionRef: CFTypeRef?
    var sizeRef: CFTypeRef?

    let positionStatus = AXUIElementCopyAttributeValue(axWindow, kAXPositionAttribute as CFString, &positionRef)
    let sizeStatus = AXUIElementCopyAttributeValue(axWindow, kAXSizeAttribute as CFString, &sizeRef)

    guard positionStatus == .success,
          sizeStatus == .success,
          let positionRaw = positionRef,
          let sizeRaw = sizeRef,
          CFGetTypeID(positionRaw) == AXValueGetTypeID(),
          CFGetTypeID(sizeRaw) == AXValueGetTypeID() else {
        return nil
    }

    let positionValue = positionRaw as! AXValue
    let sizeValue = sizeRaw as! AXValue

    guard AXValueGetType(positionValue) == .cgPoint,
          AXValueGetType(sizeValue) == .cgSize else {
        return nil
    }

    var position = CGPoint.zero
    var size = CGSize.zero

    guard AXValueGetValue(positionValue, .cgPoint, &position),
          AXValueGetValue(sizeValue, .cgSize, &size) else {
        return nil
    }

    let geometry = geometryFromRect(
        x: position.x,
        y: position.y,
        width: size.width,
        height: size.height
    )

    guard geometry.width >= 100, geometry.height >= 100 else {
        return nil
    }

    return geometry
}

private func frontmostApplicationPid() -> pid_t? {
    NSWorkspace.shared.frontmostApplication?.processIdentifier
}

private func windowStateFromAccessibilityAPI() -> WindowLookupResult? {
    let runningApps = NSWorkspace.shared.runningApplications.filter { app in
        guard let name = app.localizedName else {
            return false
        }
        return normalizedMpvName(name)
    }

    let frontmostPid = frontmostApplicationPid()
    var foundMinimizedTargetWindow = false

    for app in runningApps {
        let appElement = AXUIElementCreateApplication(app.processIdentifier)
        if !windowHasTargetSocket(app.processIdentifier) {
            continue
        }

        var windowsRef: CFTypeRef?
        let status = AXUIElementCopyAttributeValue(appElement, kAXWindowsAttribute as CFString, &windowsRef)
        guard status == .success, let windows = windowsRef as? [AXUIElement], !windows.isEmpty else {
            continue
        }

        for window in windows {
            var windowPid: pid_t = 0
            if AXUIElementGetPid(window, &windowPid) != .success {
                continue
            }

            if windowPid != app.processIdentifier {
                continue
            }

            if !windowHasTargetSocket(windowPid) {
                continue
            }

            var minimizedRef: CFTypeRef?
            let minimizedStatus = AXUIElementCopyAttributeValue(window, kAXMinimizedAttribute as CFString, &minimizedRef)
            if minimizedStatus == .success, let minimized = minimizedRef as? Bool, minimized {
                foundMinimizedTargetWindow = true
                continue
            }

            if let geometry = geometryFromAXWindow(window) {
                return .visible(
                    WindowState(
                        geometry: geometry,
                        focused: frontmostPid == windowPid
                    )
                )
            }
        }
    }

    if foundMinimizedTargetWindow {
        return .minimized
    }

    return nil
}

private func windowStateFromCoreGraphics() -> WindowState? {
    // Keep the CG fallback for environments without Accessibility permissions.
    // Use on-screen layer-0 windows to avoid off-screen helpers/shadows.
    let options: CGWindowListOption = [.optionOnScreenOnly, .excludeDesktopElements]
    let windowList = CGWindowListCopyWindowInfo(options, kCGNullWindowID) as? [[String: Any]] ?? []
    let frontmostPid = frontmostApplicationPid()

    for window in windowList {
        guard let ownerName = window[kCGWindowOwnerName as String] as? String,
              normalizedMpvName(ownerName) else {
            continue
        }

        guard let ownerPid = window[kCGWindowOwnerPID as String] as? pid_t else {
            continue
        }

        if !windowHasTargetSocket(ownerPid) {
            continue
        }

        if let layer = window[kCGWindowLayer as String] as? Int, layer != 0 {
            continue
        }
        if let alpha = window[kCGWindowAlpha as String] as? Double, alpha <= 0.01 {
            continue
        }
        if let onScreen = window[kCGWindowIsOnscreen as String] as? Int, onScreen == 0 {
            continue
        }

        guard let bounds = window[kCGWindowBounds as String] as? [String: CGFloat] else {
            continue
        }

        let geometry = geometryFromRect(
            x: bounds["X"] ?? 0,
            y: bounds["Y"] ?? 0,
            width: bounds["Width"] ?? 0,
            height: bounds["Height"] ?? 0
        )

        guard geometry.width >= 100, geometry.height >= 100 else {
            continue
        }

        return WindowState(
            geometry: geometry,
            focused: frontmostPid == ownerPid
        )
    }

    return nil
}

private let lookupResult: WindowLookupResult? = {
    if let axResult = windowStateFromAccessibilityAPI() {
        return axResult
    }
    if let cgWindow = windowStateFromCoreGraphics() {
        return .visible(cgWindow)
    }
    return nil
}()

if let result = lookupResult {
    switch result {
    case .visible(let window):
        print(
            "\(window.geometry.x),\(window.geometry.y),\(window.geometry.width),\(window.geometry.height),\(window.focused ? 1 : 0)"
        )
    case .minimized:
        print("minimized")
    }
} else {
    print("not-found")
}
