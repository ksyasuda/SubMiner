/**
 * A small right-click menu.
 *
 * Electron's native menus are a main-process affair and the browser window is
 * plain HTML, so the menu is built in the page. It is a singleton: opening one
 * closes whatever was open before, and any click, scroll, resize or Escape
 * dismisses it.
 */

export interface ContextMenuItem {
  label: string;
  onSelect: () => void;
  /** Shown greyed out and not selectable. */
  disabled?: boolean;
  /** Draws a divider above this item. */
  separated?: boolean;
}

let open: HTMLDivElement | null = null;

export function closeContextMenu(): void {
  open?.remove();
  open = null;
}

/**
 * Show `items` at the pointer. The menu is placed inside the viewport rather
 * than at the raw coordinates, so a right-click near an edge is still readable.
 */
export function showContextMenu(x: number, y: number, items: ContextMenuItem[]): void {
  closeContextMenu();
  if (items.length === 0) return;

  const menu = document.createElement('div');
  menu.className = 'context-menu';
  menu.setAttribute('role', 'menu');

  for (const item of items) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'context-menu-item';
    button.setAttribute('role', 'menuitem');
    button.textContent = item.label;
    if (item.separated) button.dataset.separated = 'true';
    if (item.disabled) {
      button.disabled = true;
    } else {
      button.addEventListener('click', () => {
        closeContextMenu();
        item.onSelect();
      });
    }
    menu.append(button);
  }

  document.body.append(menu);
  open = menu;

  const { width, height } = menu.getBoundingClientRect();
  const left = Math.max(4, Math.min(x, window.innerWidth - width - 4));
  const top = Math.max(4, Math.min(y, window.innerHeight - height - 4));
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
  menu.querySelector<HTMLButtonElement>('.context-menu-item:not(:disabled)')?.focus();
}

// Registered once: a menu that outlives the click that dismissed it is worse
// than no menu at all.
document.addEventListener('pointerdown', (event) => {
  if (open && !open.contains(event.target as Node)) closeContextMenu();
});
document.addEventListener('keydown', (event) => {
  if (!open || event.key !== 'Escape') return;
  // The detail page also closes on Escape, from a listener on this same node,
  // so stopping propagation alone would not spare it.
  event.stopImmediatePropagation();
  closeContextMenu();
});
window.addEventListener('blur', closeContextMenu);
window.addEventListener('resize', closeContextMenu);
document.addEventListener('scroll', closeContextMenu, true);
