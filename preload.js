const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('versions', {
  node: () => process.versions.node,
  chrome: () => process.versions.chrome,
  electron: () => process.versions.electron,
  ping: () => ipcRenderer.invoke('ping'),
  // we can also expose variables, not just functions
})
contextBridge.exposeInMainWorld('fs',{
  readFileSync: (filePath) => {
    return ipcRenderer.invoke("readFileSync", filePath);
  },
  readdirSyncExcel: (path) => {
    return ipcRenderer.invoke("readdirSyncExcel", path);
  },
  copyFileSync: (sourceFolder, sourceFilename, destinationFolder, destinationFilename) => {
    return ipcRenderer.invoke("copyFileSync", sourceFolder, sourceFilename, destinationFolder, destinationFilename);
  },
  createFolderSync: (parent, folderName) => {
    return ipcRenderer.invoke("createFolderSync", parent, folderName);
  },
  moveFileSync: (sourceFolder, sourceFilename, destinationFolder, destinationFilename) => {
    return ipcRenderer.invoke("moveFileSync", sourceFolder, sourceFilename, destinationFolder, destinationFilename);
  },
  runFirstExcel: (folder, fileName) => {
    return ipcRenderer.invoke("runFirstExcel", folder, fileName);
  },
  existsSync: (filepath) => {
    return ipcRenderer.invoke("existsSync", filepath);
  },
  writeFileSync: (filepath, data) => {
    return ipcRenderer.invoke("writeFileSync", filepath, data);
  },
  unlinkSync: (filepath) => {
    return ipcRenderer.invoke("unlinkSync", filepath);
  },
  unlinkSyncWithFolder: (folder, filename) => {
    return ipcRenderer.invoke("unlinkSyncWithFolder", folder, filename);
  },
  readdirSyncFolders: (path) => {
    return ipcRenderer.invoke("readdirSyncFolders", path);
  },
  rmSync: (parent, folderName) => {
    return ipcRenderer.invoke("rmSync", parent, folderName);
  },
  fileModifiedTime: (folderName, filename) => {
    return ipcRenderer.invoke("fileModifiedTime", folderName, filename);
  }
})

contextBridge.exposeInMainWorld('tc',{
  nowAsTimecode: () => {
    return ipcRenderer.invoke("nowAsTimecode");
  }
})
