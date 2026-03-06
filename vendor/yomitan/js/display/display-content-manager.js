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

import {EventListenerCollection} from '../core/event-listener-collection.js';
import {base64ToArrayBuffer} from '../data/array-buffer-util.js';

/**
 * The content manager which is used when generating HTML display content.
 */
export class DisplayContentManager {
    /**
     * Creates a new instance of the class.
     * @param {import('./display.js').Display} display The display instance that owns this object.
     */
    constructor(display) {
        /** @type {import('./display.js').Display} */
        this._display = display;
        /** @type {import('core').TokenObject} */
        this._token = {};
        /** @type {EventListenerCollection} */
        this._eventListeners = new EventListenerCollection();
        /** @type {import('display-content-manager').LoadMediaRequest[]} */
        this._loadMediaRequests = [];
        /** @type {string[]} */
        this._mediaUrls = [];
    }

    /** @type {import('display-content-manager').LoadMediaRequest[]} */
    get loadMediaRequests() {
        return this._loadMediaRequests;
    }

    /**
     * Queues loading media file from a given dictionary.
     * @param {string} path
     * @param {string} dictionary
     * @param {HTMLImageElement|HTMLCanvasElement} element
     */
    loadMedia(path, dictionary, element) {
        this._loadMediaRequests.push({path, dictionary, element});
    }

    /**
     * Unloads all media that has been loaded.
     */
    unloadAll() {
        this._token = {};

        this._eventListeners.removeAllEventListeners();

        this._loadMediaRequests = [];
        for (const mediaUrl of this._mediaUrls) {
            URL.revokeObjectURL(mediaUrl);
        }
        this._mediaUrls = [];
    }

    /**
     * Sets up attributes and events for a link element.
     * @param {HTMLAnchorElement} element The link element.
     * @param {string} href The URL.
     * @param {boolean} internal Whether or not the URL is an internal or external link.
     */
    prepareLink(element, href, internal) {
        element.href = href;
        if (!internal) {
            element.target = '_blank';
            element.rel = 'noreferrer noopener';
        }
        this._eventListeners.addEventListener(element, 'click', this._onLinkClick.bind(this));
    }

    /**
     * Execute media requests
     */
    async executeMediaRequests() {
        const token = this._token;
        for (const {path, dictionary, element} of this._loadMediaRequests) {
            try {
                const data = await this._display.application.api.getMedia([{path, dictionary}]);
                if (this._token !== token) { return; }

                const item = data[0];
                if (
                    typeof item !== 'object' ||
                    item === null ||
                    typeof item.content !== 'string' ||
                    typeof item.mediaType !== 'string'
                ) {
                    this._setMediaElementState(element, 'load-error');
                    continue;
                }

                const buffer = base64ToArrayBuffer(item.content);
                const blob = new Blob([buffer], {type: item.mediaType});
                const blobUrl = URL.createObjectURL(blob);

                if (element instanceof HTMLImageElement) {
                    this._mediaUrls.push(blobUrl);
                    element.onload = () => {
                        if (this._token !== token) { return; }
                        this._setMediaElementState(element, 'loaded');
                    };
                    element.onerror = () => {
                        if (this._token !== token) { return; }
                        this._setMediaElementState(element, 'load-error');
                    };
                    element.src = blobUrl;
                    continue;
                }

                const image = new Image();
                image.onload = () => {
                    try {
                        if (this._token !== token) { return; }
                        const context = element.getContext('2d');
                        if (context === null) {
                            this._setMediaElementState(element, 'load-error');
                            return;
                        }
                        element.width = image.naturalWidth || element.width;
                        element.height = image.naturalHeight || element.height;
                        context.drawImage(image, 0, 0, element.width, element.height);
                        this._setMediaElementState(element, 'loaded');
                    } finally {
                        URL.revokeObjectURL(blobUrl);
                    }
                };
                image.onerror = () => {
                    if (this._token === token) {
                        this._setMediaElementState(element, 'load-error');
                    }
                    URL.revokeObjectURL(blobUrl);
                };
                image.src = blobUrl;
            } catch (_e) {
                if (this._token !== token) { return; }
                this._setMediaElementState(element, 'load-error');
            }
        }
        this._loadMediaRequests = [];
    }

    /**
     * @param {string} path
     * @param {string} dictionary
     * @param {Window} window
     */
    async openMediaInTab(path, dictionary, window) {
        const data = await this._display.application.api.getMedia([{path, dictionary}]);
        const buffer = base64ToArrayBuffer(data[0].content);
        const blob = new Blob([buffer], {type: data[0].mediaType});
        const blobUrl = URL.createObjectURL(blob);
        window.open(blobUrl, '_blank')?.focus();
    }

    /**
     * @param {MouseEvent} e
     */
    _onLinkClick(e) {
        const {href} = /** @type {HTMLAnchorElement} */ (e.currentTarget);
        if (typeof href !== 'string') { return; }

        const baseUrl = new URL(location.href);
        const url = new URL(href, baseUrl);
        const internal = (url.protocol === baseUrl.protocol && url.host === baseUrl.host);
        if (!internal) { return; }

        e.preventDefault();

        /** @type {import('display').HistoryParams} */
        const params = {};
        for (const [key, value] of url.searchParams.entries()) {
            params[key] = value;
        }
        this._display.setContent({
            historyMode: 'new',
            focus: false,
            params,
            state: null,
            content: null,
        });
    }

    /**
     * @param {HTMLImageElement|HTMLCanvasElement} element
     * @param {'loaded'|'load-error'} state
     */
    _setMediaElementState(element, state) {
        const link = element.closest('.gloss-image-link');
        if (link === null) { return; }
        link.dataset.imageLoadState = state;
        if (state === 'loaded') {
            link.dataset.hasImage = 'true';
        }
    }
}
