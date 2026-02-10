#!/bin/bash
#
# SubMiner - All-in-one sentence mining overlay
# Copyright (C) 2024 sudacode
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# patch-yomitan.sh - Apply Electron compatibility patches to Yomitan
#
# This script applies the necessary patches to make Yomitan work in Electron
# after upgrading to a new version. Run this after extracting a fresh Yomitan release.
#
# Usage: ./patch-yomitan.sh [yomitan_dir]
#   yomitan_dir: Path to the Yomitan directory (default: vendor/yomitan)
#

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
YOMITAN_DIR="${1:-$SCRIPT_DIR/../vendor/yomitan}"

if [ ! -d "$YOMITAN_DIR" ]; then
    echo "Error: Yomitan directory not found: $YOMITAN_DIR"
    exit 1
fi

echo "Patching Yomitan in: $YOMITAN_DIR"

PERMISSIONS_UTIL="$YOMITAN_DIR/js/data/permissions-util.js"

if [ ! -f "$PERMISSIONS_UTIL" ]; then
    echo "Error: permissions-util.js not found at $PERMISSIONS_UTIL"
    exit 1
fi

echo "Patching permissions-util.js..."

if grep -q "Electron workaround" "$PERMISSIONS_UTIL"; then
    echo "  - Already patched, skipping"
else
    cat > "$PERMISSIONS_UTIL.tmp" << 'PATCH_EOF'
/*
 * Copyright (C) 2023-2025  Yomitan Authors
 * Copyright (C) 2021-2022  Yomichan Authors
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import {getFieldMarkers} from './anki-util.js';

/**
 * This function returns whether an Anki field marker might require clipboard permissions.
 * This is speculative and may not guarantee that the field marker actually does require the permission,
 * as the custom handlebars template is not deeply inspected.
 * @param {string} marker
 * @returns {boolean}
 */
function ankiFieldMarkerMayUseClipboard(marker) {
    switch (marker) {
        case 'clipboard-image':
        case 'clipboard-text':
            return true;
        default:
            return false;
    }
}

/**
 * @param {chrome.permissions.Permissions} permissions
 * @returns {Promise<boolean>}
 */
export function hasPermissions(permissions) {
    return new Promise((resolve, reject) => {
        chrome.permissions.contains(permissions, (result) => {
            const e = chrome.runtime.lastError;
            if (e) {
                reject(new Error(e.message));
            } else {
                resolve(result);
            }
        });
    });
}

/**
 * @param {chrome.permissions.Permissions} permissions
 * @param {boolean} shouldHave
 * @returns {Promise<boolean>}
 */
export function setPermissionsGranted(permissions, shouldHave) {
    return (
        shouldHave ?
        new Promise((resolve, reject) => {
            chrome.permissions.request(permissions, (result) => {
                const e = chrome.runtime.lastError;
                if (e) {
                    reject(new Error(e.message));
                } else {
                    resolve(result);
                }
            });
        }) :
        new Promise((resolve, reject) => {
            chrome.permissions.remove(permissions, (result) => {
                const e = chrome.runtime.lastError;
                if (e) {
                    reject(new Error(e.message));
                } else {
                    resolve(!result);
                }
            });
        })
    );
}

/**
 * @returns {Promise<chrome.permissions.Permissions>}
 */
export function getAllPermissions() {
    // Electron workaround - chrome.permissions.getAll() not available
    return Promise.resolve({
        origins: ["<all_urls>"],
        permissions: ["clipboardWrite", "storage", "unlimitedStorage", "scripting", "contextMenus"]
    });
}

/**
 * @param {string} fieldValue
 * @returns {string[]}
 */
export function getRequiredPermissionsForAnkiFieldValue(fieldValue) {
    const markers = getFieldMarkers(fieldValue);
    for (const marker of markers) {
        if (ankiFieldMarkerMayUseClipboard(marker)) {
            return ['clipboardRead'];
        }
    }
    return [];
}

/**
 * @param {chrome.permissions.Permissions} permissions
 * @param {import('settings').ProfileOptions} options
 * @returns {boolean}
 */
export function hasRequiredPermissionsForOptions(permissions, options) {
    const permissionsSet = new Set(permissions.permissions);

    if (!permissionsSet.has('nativeMessaging') && (options.parsing.enableMecabParser || options.general.enableYomitanApi)) {
        return false;
    }

    if (!permissionsSet.has('clipboardRead')) {
        if (options.clipboard.enableBackgroundMonitor || options.clipboard.enableSearchPageMonitor) {
            return false;
        }
        const fieldsList = options.anki.cardFormats.map((cardFormat) => cardFormat.fields);

        for (const fields of fieldsList) {
            for (const {value: fieldValue} of Object.values(fields)) {
                const markers = getFieldMarkers(fieldValue);
                for (const marker of markers) {
                    if (ankiFieldMarkerMayUseClipboard(marker)) {
                        return false;
                    }
                }
            }
        }
    }

    return true;
}
PATCH_EOF

    mv "$PERMISSIONS_UTIL.tmp" "$PERMISSIONS_UTIL"
    echo "  - Patched successfully"
fi

OPTIONS_SCHEMA="$YOMITAN_DIR/data/schemas/options-schema.json"

if [ ! -f "$OPTIONS_SCHEMA" ]; then
    echo "Error: options-schema.json not found at $OPTIONS_SCHEMA"
    exit 1
fi

echo "Patching options-schema.json..."

if grep -q '"selectText".*"default": true' "$OPTIONS_SCHEMA"; then
    sed -i '/"selectText": {/,/"default":/{s/"default": true/"default": false/}' "$OPTIONS_SCHEMA"
    echo "  - Changed selectText default to false"
elif grep -q '"selectText".*"default": false' "$OPTIONS_SCHEMA"; then
    echo "  - selectText already set to false, skipping"
else
    echo "  - Warning: Could not find selectText setting"
fi

if grep -q '"layoutAwareScan".*"default": true' "$OPTIONS_SCHEMA"; then
    sed -i '/"layoutAwareScan": {/,/"default":/{s/"default": true/"default": false/}' "$OPTIONS_SCHEMA"
    echo "  - Changed layoutAwareScan default to false"
elif grep -q '"layoutAwareScan".*"default": false' "$OPTIONS_SCHEMA"; then
    echo "  - layoutAwareScan already set to false, skipping"
else
    echo "  - Warning: Could not find layoutAwareScan setting"
fi

POPUP_JS="$YOMITAN_DIR/js/app/popup.js"

if [ ! -f "$POPUP_JS" ]; then
    echo "Error: popup.js not found at $POPUP_JS"
    exit 1
fi

echo "Patching popup.js..."

if grep -q "yomitan-popup-shown" "$POPUP_JS"; then
    echo "  - Already patched, skipping"
else
    # Add the visibility event dispatch after the existing _onVisibleChange code
    # We need to add it after: void this._invokeSafe('displayVisibilityChanged', {value});
    sed -i "/void this._invokeSafe('displayVisibilityChanged', {value});/a\\
\\
        // Dispatch custom events for popup visibility (Electron integration)\\
        if (value) {\\
            window.dispatchEvent(new CustomEvent('yomitan-popup-shown'));\\
        } else {\\
            window.dispatchEvent(new CustomEvent('yomitan-popup-hidden'));\\
        }" "$POPUP_JS"
    echo "  - Added visibility events"
fi

echo ""
echo "Yomitan patching complete!"
echo ""
echo "Changes applied:"
echo "  1. permissions-util.js: Hardcoded permissions (Electron workaround)"
echo "  2. options-schema.json: selectText=false, layoutAwareScan=false"
echo "  3. popup.js: Added yomitan-popup-shown/hidden events"
echo ""
echo "To verify: Run 'pnpm start --dev' and check for 'Yomitan extension loaded successfully'"
