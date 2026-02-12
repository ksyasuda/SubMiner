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

private struct WindowGeometry {
    let x: Int
    let y: Int
    let width: Int
    let height: Int
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

private func geometryFromAccessibilityAPI() -> WindowGeometry? {
    let runningApps = NSWorkspace.shared.runningApplications.filter { app in
        guard let name = app.localizedName else {
            return false
        }
        return normalizedMpvName(name)
    }

    for app in runningApps {
        let appElement = AXUIElementCreateApplication(app.processIdentifier)
        var windowsRef: CFTypeRef?
        let status = AXUIElementCopyAttributeValue(appElement, kAXWindowsAttribute as CFString, &windowsRef)
        guard status == .success, let windows = windowsRef as? [AXUIElement], !windows.isEmpty else {
            continue
        }

        for window in windows {
            var minimizedRef: CFTypeRef?
            let minimizedStatus = AXUIElementCopyAttributeValue(window, kAXMinimizedAttribute as CFString, &minimizedRef)
            if minimizedStatus == .success, let minimized = minimizedRef as? Bool, minimized {
                continue
            }

            if let geometry = geometryFromAXWindow(window) {
                return geometry
            }
        }
    }

    return nil
}

private func geometryFromCoreGraphics() -> WindowGeometry? {
    // Keep the CG fallback for environments without Accessibility permissions.
    // Use on-screen layer-0 windows to avoid off-screen helpers/shadows.
    let options: CGWindowListOption = [.optionOnScreenOnly, .excludeDesktopElements]
    let windowList = CGWindowListCopyWindowInfo(options, kCGNullWindowID) as? [[String: Any]] ?? []

    for window in windowList {
        guard let ownerName = window[kCGWindowOwnerName as String] as? String,
              normalizedMpvName(ownerName) else {
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

        return geometry
    }

    return nil
}

if let window = geometryFromAccessibilityAPI() ?? geometryFromCoreGraphics() {
    print("\(window.x),\(window.y),\(window.width),\(window.height)")
} else {
    print("not-found")
}
