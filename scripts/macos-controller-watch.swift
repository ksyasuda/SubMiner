#!/usr/bin/env swift
//
//  macos-controller-watch.swift
//  SubMiner - Report controller button presses on macOS while SubMiner is in the background
//
//  macOS only delivers GameController input to the frontmost app, so the overlay sees no
//  presses while mpv is in front. This helper opts in to background events and prints one
//  line per press so SubMiner can focus the overlay and replay the press there.
//
//  Usage: macos-controller-watch (long-running; exits when stdin closes)
//  Output: "press <index> <connected pads>" per button press, using the W3C standard gamepad
//  button indices. SubMiner only replays a press when one pad is connected: nothing identifies
//  a pad to both GameController and Chromium, so with several it could hit the wrong one.
//

import Cocoa
import Foundation
import GameController

private func emit(_ line: String) {
    FileHandle.standardOutput.write(Data((line + "\n").utf8))
}

/// Standard gamepad mapping (https://w3c.github.io/gamepad/#remapping), matching the
/// indices Chromium reports for the same pad.
private func standardButtons(_ pad: GCExtendedGamepad) -> [(Int, GCControllerButtonInput?)] {
    return [
        (0, pad.buttonA),
        (1, pad.buttonB),
        (2, pad.buttonX),
        (3, pad.buttonY),
        (4, pad.leftShoulder),
        (5, pad.rightShoulder),
        (6, pad.leftTrigger),
        (7, pad.rightTrigger),
        (8, pad.buttonOptions),
        (9, pad.buttonMenu),
        (10, pad.leftThumbstickButton),
        (11, pad.rightThumbstickButton),
        (12, pad.dpad.up),
        (13, pad.dpad.down),
        (14, pad.dpad.left),
        (15, pad.dpad.right),
        (16, pad.buttonHome),
    ]
}

private func connectedPadCount() -> Int {
    return GCController.controllers().filter { $0.extendedGamepad != nil }.count
}

private func watch(_ controller: GCController) {
    guard let pad = controller.extendedGamepad else { return }
    for (index, button) in standardButtons(pad) {
        button?.pressedChangedHandler = { _, _, pressed in
            if pressed {
                emit("press \(index) \(connectedPadCount())")
            }
        }
    }
}

// Exit with SubMiner: it holds our stdin open for as long as it runs.
private func exitWhenStdinCloses() {
    Thread.detachNewThread {
        while !FileHandle.standardInput.availableData.isEmpty {}
        exit(0)
    }
}

let app = NSApplication.shared
app.setActivationPolicy(.prohibited)
GCController.shouldMonitorBackgroundEvents = true

NotificationCenter.default.addObserver(
    forName: .GCControllerDidConnect,
    object: nil,
    queue: .main
) { notification in
    if let controller = notification.object as? GCController {
        watch(controller)
    }
}
GCController.controllers().forEach(watch)

exitWhenStdinCloses()
app.run()
