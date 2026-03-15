import { contextBridge, ipcRenderer } from 'electron';
import { IPC_CHANNELS } from './shared/ipc/contracts';

const statsAPI = {
  getOverview: (): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetOverview),

  getDailyRollups: (limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetDailyRollups, limit),

  getMonthlyRollups: (limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMonthlyRollups, limit),

  getSessions: (limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetSessions, limit),

  getSessionTimeline: (sessionId: number, limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetSessionTimeline, sessionId, limit),

  getSessionEvents: (sessionId: number, limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetSessionEvents, sessionId, limit),

  getVocabulary: (limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetVocabulary, limit),

  getKanji: (limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetKanji, limit),

  getMediaLibrary: (): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaLibrary),

  getMediaDetail: (videoId: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaDetail, videoId),

  getMediaSessions: (videoId: number, limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaSessions, videoId, limit),

  getMediaDailyRollups: (videoId: number, limit?: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaDailyRollups, videoId, limit),

  getMediaCover: (videoId: number): Promise<unknown> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.statsGetMediaCover, videoId),

  hideOverlay: (): void => {
    ipcRenderer.send(IPC_CHANNELS.command.toggleStatsOverlay);
  },
};

contextBridge.exposeInMainWorld('electronAPI', { stats: statsAPI });
