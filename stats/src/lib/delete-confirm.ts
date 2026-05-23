type NativeDialogBridge = {
  electronAPI?: {
    stats?: {
      confirmNativeDialog?: (message: string) => boolean;
      beginNativeDialog?: () => void;
      endNativeDialog?: () => void;
    };
  };
};

type DeleteConfirmPresenter = (message: string) => boolean | Promise<boolean>;

let deleteConfirmPresenter: DeleteConfirmPresenter | null = null;

export function setDeleteConfirmPresenter(presenter: DeleteConfirmPresenter): () => void {
  deleteConfirmPresenter = presenter;
  return () => {
    if (deleteConfirmPresenter === presenter) {
      deleteConfirmPresenter = null;
    }
  };
}

async function confirmWithStatsNativeDialogLayer(message: string): Promise<boolean> {
  if (deleteConfirmPresenter) {
    return deleteConfirmPresenter(message);
  }

  const statsApi = (globalThis as typeof globalThis & NativeDialogBridge).electronAPI?.stats;
  if (statsApi?.confirmNativeDialog) {
    return statsApi.confirmNativeDialog(message);
  }

  statsApi?.beginNativeDialog?.();
  try {
    return globalThis.confirm(message);
  } finally {
    statsApi?.endNativeDialog?.();
  }
}

export function confirmSessionDelete(): Promise<boolean> {
  return confirmWithStatsNativeDialogLayer('Delete this session and all associated data?');
}

export function confirmDayGroupDelete(dayLabel: string, count: number): Promise<boolean> {
  return confirmWithStatsNativeDialogLayer(
    `Delete all ${count} session${count === 1 ? '' : 's'} from ${dayLabel} and all associated data?`,
  );
}

export function confirmAnimeGroupDelete(title: string, count: number): Promise<boolean> {
  return confirmWithStatsNativeDialogLayer(
    `Delete all ${count} session${count === 1 ? '' : 's'} for "${title}" and all associated data?`,
  );
}

export function confirmEpisodeDelete(title: string): Promise<boolean> {
  return confirmWithStatsNativeDialogLayer(`Delete "${title}" and all its sessions?`);
}

export function confirmBucketDelete(title: string, count: number): Promise<boolean> {
  if (count === 1) {
    return confirmWithStatsNativeDialogLayer(
      `Delete this session of "${title}" from this day and all associated data?`,
    );
  }
  return confirmWithStatsNativeDialogLayer(
    `Delete all ${count} sessions of "${title}" from this day and all associated data?`,
  );
}
