import { contextBridge, ipcRenderer } from 'electron';
import { createAnimeBrowserAPI } from './preload-anime-browser-api';

contextBridge.exposeInMainWorld('animeBrowserAPI', createAnimeBrowserAPI(ipcRenderer));
