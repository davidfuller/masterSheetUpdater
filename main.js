const { app, BrowserWindow, ipcMain, screen } = require('electron')
const path = require('node:path')
const fs = require('fs');
const tc = require("./timecode.js")

const createWindow = () => {
  const primaryDisplay = screen.getPrimaryDisplay()
  const screenWidth = primaryDisplay.workAreaSize.width
  const win = new BrowserWindow({
    width: 800,
    height: 1200,
    x: screenWidth - 800,
    y: 50,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })

  win.loadFile('index.html');
  win.setBackgroundColor('#f0a800');
}

app.whenReady().then(() => {
  ipcMain.handle("readFileSync", (event, filePath) => {
    let doIt = false;
    try {
      doIt = fs.existsSync(filePath);
    } catch (err) {
      console.log('Error in file exists:', filePath, err);
    }
    if (doIt){
      return fs.readFileSync(filePath, 'utf8');
    }
  });
  
  ipcMain.handle('readdirSyncExcel', (event, path) => {
    let doIt = false;
    try {
      doIt = fs.existsSync(path);
    } catch (err) {
      console.log('Error in file exists:', path, err);
    }
    if (doIt){
      let filenames = fs.readdirSync(path, {withFileTypes: false});
      return filenames.filter(isExcel).sort();
    } else {
      return [];
    }
  });

  ipcMain.handle("copyFileSync", (event, sourceFolder, sourceFilename, destinationFolder, destinationFilename) => {
    let doIt = false;
    let doItDest = false;
    const source = path.join(sourceFolder,sourceFilename);
    const destFolder = path.join(destinationFolder);
    try {
      doIt = fs.existsSync(source)
      doItDest = fs.existsSync(destFolder);
    } catch (err) {
      console.log('Error in file exists:', path, err);
    }
    if (doIt & doItDest){
      const destination = path.join(destinationFolder, destinationFilename);
      fs.copyFileSync(source, destination);
    }
  });


  ipcMain.handle("createFolderSync", (event, parent, folderName) => {
    const newFolderName = path.join(parent, folderName);
    if (!fs.existsSync(newFolderName)){
      fs.mkdirSync(newFolderName);
    }
    return newFolderName;
  });
  
  ipcMain.handle("moveFileSync", (event, sourceFolder, sourceFilename, destinationFolder, destinationFilename) => {
    let doIt = false;
    let doItDest = false;
    const source = path.join(sourceFolder,sourceFilename);
    const destination = path.join(destinationFolder, destinationFilename);
    try {
      doIt = fs.existsSync(source)
      doItDest = fs.existsSync(destination);
    } catch (err) {
      console.log('Error in file exists:', path, err);
    }
    if (doIt & doItDest){
      fs.renameSync(source, destination);
    }
  });

  ipcMain.handle("runFirstExcel", (event, folder, filename) => {
    let theBat = path.join(__dirname,'startExcel.bat')
    let excelFile = path.join(folder, filename)
    let command = theBat + ' "' + excelFile + '"'
    console.log(command)
    runWindowsBat(command);
  });
  ipcMain.handle("nowAsTimecode", (event) =>{
    return tc.nowAsTimecode();
  });
  ipcMain.handle("existsSync", (event, filepath) => {
    return fs.existsSync(filepath);
  });
  ipcMain.handle("writeFileSync", (event, filepath, data) => {
    fs.writeFileSync(filepath, data, "utf8");
  });
  ipcMain.handle("unlinkSync", (event, filepath) => {
    fs.unlinkSync(filepath);
  });
  ipcMain.handle("unlinkSyncWithFolder", (event, folder, filename) => {
    let filepath = path.join(folder, filename)
    fs.unlinkSync(filepath);
  });

  ipcMain.handle('readdirSyncFolders', (event, path) => {
    let filenames = fs.readdirSync(path, {withFileTypes: true});
    return filenames.filter(filename => filename.isDirectory()).map(filename => filename.name);
  });

  ipcMain.handle("rmSync", (event, parent, folderName) => {
    const folder = path.join(parent, folderName);
    return fs.rmSync(folder, {recursive: true});
  });

  ipcMain.handle("fileModifiedTime", (event, folderName, filename) => {
    const fileToTest = path.join(folderName, filename);
    return fs.statSync(fileToTest).mtimeMs;
  });



  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

function isExcel(filename){
  return (path.extname(filename) == ".xlsm"  && !filename.startsWith("~$"));
}

function runWindowsBat(path){
  const { exec } = require('child_process');
  exec(path, (err, stdout, stderr) => {
    if (err) {
      console.error(err);
      return;
    }
    console.log(stdout);
  });
}