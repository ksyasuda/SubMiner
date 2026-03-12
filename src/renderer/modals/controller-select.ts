import type { ModalStateReader, RendererContext } from '../context';

function clampSelectedIndex(ctx: RendererContext): void {
  if (ctx.state.connectedGamepads.length === 0) {
    ctx.state.controllerDeviceSelectedIndex = 0;
    return;
  }

  ctx.state.controllerDeviceSelectedIndex = Math.min(
    Math.max(ctx.state.controllerDeviceSelectedIndex, 0),
    ctx.state.connectedGamepads.length - 1,
  );
}

export function createControllerSelectModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let selectedControllerId: string | null = null;
  let lastRenderedDevicesKey = '';
  let lastRenderedActiveGamepadId: string | null = null;
  let lastRenderedPreferredId = '';

  function getDevicesKey(): string {
    return ctx.state.connectedGamepads
      .map((device) => `${device.id}|${device.index}|${device.mapping}|${device.connected}`)
      .join('||');
  }

  function syncSelectedControllerId(): void {
    const selected = ctx.state.connectedGamepads[ctx.state.controllerDeviceSelectedIndex];
    selectedControllerId = selected?.id ?? null;
  }

  function syncSelectedIndexToCurrentController(): void {
    const preferredId = ctx.state.controllerConfig?.preferredGamepadId ?? '';
    const activeIndex = ctx.state.connectedGamepads.findIndex(
      (device) => device.id === ctx.state.activeGamepadId,
    );
    if (activeIndex >= 0) {
      ctx.state.controllerDeviceSelectedIndex = activeIndex;
      syncSelectedControllerId();
      return;
    }
    const preferredIndex = ctx.state.connectedGamepads.findIndex((device) => device.id === preferredId);
    if (preferredIndex >= 0) {
      ctx.state.controllerDeviceSelectedIndex = preferredIndex;
      syncSelectedControllerId();
      return;
    }
    clampSelectedIndex(ctx);
    syncSelectedControllerId();
  }

  function setStatus(message: string, isError = false): void {
    ctx.dom.controllerSelectStatus.textContent = message;
    ctx.dom.controllerSelectStatus.classList.toggle('error', isError);
  }

  function renderList(): void {
    ctx.dom.controllerSelectList.innerHTML = '';
    clampSelectedIndex(ctx);

    const preferredId = ctx.state.controllerConfig?.preferredGamepadId ?? '';
    ctx.state.connectedGamepads.forEach((device, index) => {
      const li = document.createElement('li');
      li.className = 'runtime-options-list-entry';

      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'runtime-options-item runtime-options-item-button';
      button.classList.toggle('active', index === ctx.state.controllerDeviceSelectedIndex);

      const label = document.createElement('div');
      label.className = 'runtime-options-label';
      label.textContent = device.id || `Gamepad ${device.index}`;

      const meta = document.createElement('div');
      meta.className = 'runtime-options-value';
      const tags = [
        `Index ${device.index}`,
        device.mapping || 'unknown mapping',
        device.id === ctx.state.activeGamepadId ? 'active' : null,
        device.id === preferredId ? 'saved' : null,
      ].filter(Boolean);
      meta.textContent = tags.join(' · ');

      button.appendChild(label);
      button.appendChild(meta);
      button.addEventListener('click', () => {
        ctx.state.controllerDeviceSelectedIndex = index;
        syncSelectedControllerId();
        renderList();
      });
      button.addEventListener('dblclick', () => {
        ctx.state.controllerDeviceSelectedIndex = index;
        syncSelectedControllerId();
        void saveSelectedController();
      });
      li.appendChild(button);

      ctx.dom.controllerSelectList.appendChild(li);
    });

    lastRenderedDevicesKey = getDevicesKey();
    lastRenderedActiveGamepadId = ctx.state.activeGamepadId;
    lastRenderedPreferredId = preferredId;
  }

  function updateDevices(): void {
    if (!ctx.state.controllerSelectModalOpen) return;
    if (selectedControllerId) {
      const preservedIndex = ctx.state.connectedGamepads.findIndex(
        (device) => device.id === selectedControllerId,
      );
      if (preservedIndex >= 0) {
        ctx.state.controllerDeviceSelectedIndex = preservedIndex;
      } else {
        syncSelectedIndexToCurrentController();
      }
    } else {
      syncSelectedIndexToCurrentController();
    }

    const preferredId = ctx.state.controllerConfig?.preferredGamepadId ?? '';
    const shouldRender =
      getDevicesKey() !== lastRenderedDevicesKey ||
      ctx.state.activeGamepadId !== lastRenderedActiveGamepadId ||
      preferredId !== lastRenderedPreferredId;
    if (shouldRender) {
      renderList();
    }

    if (ctx.state.connectedGamepads.length === 0) {
      setStatus('No controllers detected.');
      return;
    }
    const currentStatus = ctx.dom.controllerSelectStatus.textContent.trim();
    if (
      currentStatus !== 'No controller selected.' &&
      !currentStatus.startsWith('Saved preferred controller:')
    ) {
      setStatus('Select a controller to save as preferred.');
    }
  }

  async function saveSelectedController(): Promise<void> {
    const selected = ctx.state.connectedGamepads[ctx.state.controllerDeviceSelectedIndex];
    if (!selected) {
      setStatus('No controller selected.', true);
      return;
    }

    await window.electronAPI.saveControllerPreference({
      preferredGamepadId: selected.id,
      preferredGamepadLabel: selected.id,
    });

    if (ctx.state.controllerConfig) {
      ctx.state.controllerConfig.preferredGamepadId = selected.id;
      ctx.state.controllerConfig.preferredGamepadLabel = selected.id;
    }
    syncSelectedControllerId();
    renderList();
    setStatus(`Saved preferred controller: ${selected.id || `Gamepad ${selected.index}`}`);
  }

  function openControllerSelectModal(): void {
    ctx.state.controllerSelectModalOpen = true;
    syncSelectedIndexToCurrentController();
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.controllerSelectModal.classList.remove('hidden');
    ctx.dom.controllerSelectModal.setAttribute('aria-hidden', 'false');
    window.focus();
    ctx.dom.overlay.focus({ preventScroll: true });
    renderList();
    if (ctx.state.connectedGamepads.length === 0) {
      setStatus('No controllers detected.');
    } else {
      setStatus('Select a controller to save as preferred.');
    }
  }

  function closeControllerSelectModal(): void {
    if (!ctx.state.controllerSelectModalOpen) return;
    ctx.state.controllerSelectModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.controllerSelectModal.classList.add('hidden');
    ctx.dom.controllerSelectModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('controller-select');
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function handleControllerSelectKeydown(event: KeyboardEvent): boolean {
    if (event.key === 'Escape') {
      event.preventDefault();
      closeControllerSelectModal();
      return true;
    }

    if (event.key === 'ArrowDown' || event.key === 'j' || event.key === 'J') {
      event.preventDefault();
      if (ctx.state.connectedGamepads.length > 0) {
        ctx.state.controllerDeviceSelectedIndex = Math.min(
          ctx.state.connectedGamepads.length - 1,
          ctx.state.controllerDeviceSelectedIndex + 1,
        );
        syncSelectedControllerId();
        renderList();
      }
      return true;
    }

    if (event.key === 'ArrowUp' || event.key === 'k' || event.key === 'K') {
      event.preventDefault();
      if (ctx.state.connectedGamepads.length > 0) {
        ctx.state.controllerDeviceSelectedIndex = Math.max(
          0,
          ctx.state.controllerDeviceSelectedIndex - 1,
        );
        syncSelectedControllerId();
        renderList();
      }
      return true;
    }

    if (event.key === 'Enter') {
      event.preventDefault();
      void saveSelectedController();
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.controllerSelectClose.addEventListener('click', () => {
      closeControllerSelectModal();
    });
    ctx.dom.controllerSelectSave.addEventListener('click', () => {
      void saveSelectedController();
    });
  }

  return {
    openControllerSelectModal,
    closeControllerSelectModal,
    handleControllerSelectKeydown,
    updateDevices,
    wireDomEvents,
  };
}
