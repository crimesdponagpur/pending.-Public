// पानाला फक्त फोल्डर निवडण्याची सोय (Mac ची खिडकी)
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('shodh', {
  pickFolder: () => ipcRenderer.invoke('pick-folder')
});
