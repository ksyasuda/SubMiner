import { useCallback, useEffect, useRef, useState } from 'react';
import { useTranslation } from '../../i18n';
import { setDeleteConfirmPresenter } from '../../lib/delete-confirm';

interface PendingDeleteConfirm {
  message: string;
  resolve: (confirmed: boolean) => void;
}

export function DeleteConfirmDialog() {
  const { t } = useTranslation();
  const [pendingConfirm, setPendingConfirm] = useState<PendingDeleteConfirm | null>(null);
  const pendingRef = useRef<PendingDeleteConfirm | null>(null);
  const cancelButtonRef = useRef<HTMLButtonElement>(null);

  const finish = useCallback((confirmed: boolean) => {
    const pending = pendingRef.current;
    pendingRef.current = null;
    setPendingConfirm(null);
    pending?.resolve(confirmed);
  }, []);

  useEffect(() => {
    return setDeleteConfirmPresenter(
      (message) =>
        new Promise<boolean>((resolve) => {
          pendingRef.current?.resolve(false);
          const next = { message, resolve };
          pendingRef.current = next;
          setPendingConfirm(next);
        }),
    );
  }, []);

  useEffect(() => {
    if (!pendingConfirm) return;
    cancelButtonRef.current?.focus();
  }, [pendingConfirm]);

  useEffect(() => {
    if (!pendingConfirm) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      finish(false);
    };
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, [finish, pendingConfirm]);

  useEffect(() => {
    return () => {
      pendingRef.current?.resolve(false);
      pendingRef.current = null;
    };
  }, []);

  if (!pendingConfirm) return null;

  return (
    <div className="fixed inset-0 z-[2147483647] flex items-center justify-center bg-ctp-crust/55 p-4 backdrop-blur-sm">
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="delete-confirm-title"
        className="w-full max-w-md rounded-lg border border-ctp-surface1 bg-ctp-mantle shadow-2xl"
      >
        <div className="border-b border-ctp-surface1 px-4 py-3">
          <h2 id="delete-confirm-title" className="text-sm font-semibold text-ctp-text">
            {t('stats.delete.confirm')}
          </h2>
        </div>
        <div className="px-4 py-4 text-sm leading-6 text-ctp-subtext0">
          {pendingConfirm.message}
        </div>
        <div className="grid grid-cols-2 border-t border-ctp-surface1">
          <button
            ref={cancelButtonRef}
            type="button"
            onClick={() => finish(false)}
            className="border-r border-ctp-surface1 px-4 py-3 text-sm text-ctp-subtext0 transition-colors hover:bg-ctp-surface0 hover:text-ctp-text focus:outline-none focus:bg-ctp-surface0 focus:text-ctp-text"
          >
            {t('stats.delete.cancel')}
          </button>
          <button
            type="button"
            onClick={() => finish(true)}
            className="px-4 py-3 text-sm font-semibold text-ctp-red transition-colors hover:bg-ctp-surface0 focus:outline-none focus:bg-ctp-surface0"
          >
            {t('stats.delete.delete')}
          </button>
        </div>
      </div>
    </div>
  );
}
