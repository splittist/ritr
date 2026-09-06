import { contextBridge, ipcRenderer } from 'electron';
import type { DesktopApi } from './protocol';

const api: DesktopApi = {
  snapshot: () => ipcRenderer.invoke('ritr:snapshot'),
  open: () => ipcRenderer.invoke('ritr:open'),
  previewEdit: (edit) => ipcRenderer.invoke('ritr:previewEdit', edit),
  previewReplace: (query, replacement, caseSensitive) =>
    ipcRenderer.invoke('ritr:previewReplace', query, replacement, caseSensitive),
  search: (query, caseSensitive) => ipcRenderer.invoke('ritr:search', query, caseSensitive),
  commit: (id) => ipcRenderer.invoke('ritr:commit', id),
  undo: () => ipcRenderer.invoke('ritr:undo'),
  redo: () => ipcRenderer.invoke('ritr:redo'),
  save: (id) => ipcRenderer.invoke('ritr:save', id),
  report: (id) => ipcRenderer.invoke('ritr:report', id),
};
contextBridge.exposeInMainWorld('ritr', api);
