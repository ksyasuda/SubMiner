/*
 * Copyright (C) 2023-2025  Yomitan Authors
 * Copyright (C) 2020-2022  Yomichan Authors
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

import {Application} from '../application.js';
import {DocumentFocusController} from '../dom/document-focus-controller.js';
import {HotkeyHandler} from '../input/hotkey-handler.js';
import {DisplayAnki} from './display-anki.js';
import {DisplayAudio} from './display-audio.js';
import {DisplayProfileSelection} from './display-profile-selection.js';
import {DisplayResizer} from './display-resizer.js';
import {Display} from './display.js';

await Application.main(true, async (application) => {
    const documentFocusController = new DocumentFocusController();
    documentFocusController.prepare();

    const hotkeyHandler = new HotkeyHandler();
    hotkeyHandler.prepare(application.crossFrame);

    const display = new Display(application, 'popup', documentFocusController, hotkeyHandler);
    await display.prepare();

    const displayAudio = new DisplayAudio(display);
    displayAudio.prepare();

    const displayAnki = new DisplayAnki(display, displayAudio);
    displayAnki.prepare();

    const displayProfileSelection = new DisplayProfileSelection(display);
    void displayProfileSelection.prepare();

    const displayResizer = new DisplayResizer(display);
    displayResizer.prepare();

    document.addEventListener('keydown', (event) => {
        if (event.defaultPrevented) { return; }
        if (event.ctrlKey || event.metaKey || event.altKey) { return; }

        const target = /** @type {?Element} */ (event.target instanceof Element ? event.target : null);
        if (target !== null) {
            if (target.closest('input, textarea, select, [contenteditable="true"]')) {
                return;
            }
        }

        const code = event.code;
        const isPopupScrollKey =
            code === 'KeyJ' ||
            code === 'KeyK' ||
            code === 'ArrowDown' ||
            code === 'ArrowUp';
        if (isPopupScrollKey) {
            const scanningOptions = display.getOptions()?.scanning;
            const scale = Number.isFinite(scanningOptions?.reducedMotionScrollingScale)
                ? scanningOptions.reducedMotionScrollingScale
                : 1;
            display._scrollByPopupHeight(
                code === 'KeyJ' || code === 'ArrowDown' ? 1 : -1,
                scale,
            );
            event.preventDefault();
            return;
        }

        if (code === 'KeyM') {
            if (event.repeat) { return; }
            displayAnki._hotkeySaveAnkiNoteForSelectedEntry('0');
            event.preventDefault();
            return;
        }

        if (code === 'KeyP') {
            if (event.repeat) { return; }
            void displayAudio.playAudio(display.selectedIndex, 0);
            event.preventDefault();
            return;
        }

        if (code === 'BracketLeft' || code === 'BracketRight') {
            if (event.repeat) { return; }
            displayAudio._onMessageCycleAudioSource({direction: code === 'BracketLeft' ? 1 : -1});
            event.preventDefault();
        }
    });

    document.addEventListener('subminer-display-mine-selected', () => {
        displayAnki._hotkeySaveAnkiNoteForSelectedEntry('0');
    });

    display.initializeState();

    document.documentElement.dataset.loaded = 'true';
});
