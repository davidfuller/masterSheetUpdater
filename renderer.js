const information = document.getElementById('info');
const currentExcelInfo = document.getElementById('currentExcel');
const currentTime = document.getElementById('currentTime')
const btnDirOneDriveToProcess = document.getElementById('btnRead');
const btnTidyUpProcessed = document.getElementById("btnTidyProcessed");
const btnTidyUpOriginals = document.getElementById("btnTidyOriginals");
const btnRunFirstExcel = document.getElementById("btnRunFirstExcel");
const btnFileProcessed = document.getElementById("btnFileProcessed")
const statusList = document.getElementById('txtStatus');
const btnCopyOriginals = document.getElementById("btnCopyOriginals");
const btnCopyProcessed = document.getElementById("btnCopyProcessed");
const btnCopySpecialsSheets = document.getElementById("btnCopySpecialsSheets");
const btnTest = document.getElementById("btnTest");
const btnDeleteToProcessOneDrive = document.getElementById("btnDeleteToProcessOneDrive");
const btnRunSpecialExcel = document.getElementById("btnRunSpecialExcel");
const btnDeleteOldFolders = document.getElementById("btnDeleteOldFolders");
const scheduleColumn01 = document.getElementById("column1")
const scheduleColumn02 = document.getElementById("column2")
const scheduleColumn03 = document.getElementById("column3")
const btnToggleSchedule = document.getElementById("btnToggleSchedule");
const btnToggleButtons = document.getElementById("btnToggleButtons");
const divSchedule = document.getElementById("schedule");
const divButtons = document.getElementById("theButtons");
const divMainStatus = document.getElementById("mainStatus");

const msgCopyToProcessOneDriveToDropbox = "Copying 'To Process' files from OneDrive to Dropbox";

const eventStartProcessingSpecialSheets = "startProcessingSpecialSheets";
const eventStartExcelProcessing = "startExcelProcessing";

let currentExcel;
let currentSpecialsExcel;
let finishedCounter = 0;
let finishedSpecialsCounter = 0;
let doRunning = false;
let doSpecials = false;
let currentEvents = [];
let doingMainLoop = false;
let pollTimes = [];
let todaysSheets = [];

information.innerText = `This app is using Chrome (v${versions.chrome()}), Node.js (v${versions.node()}), and Electron (v${versions.electron()})`
setInterval(getTime, 40 );

setInterval(isCurrentExcelProcessed, 120000);
setInterval(isCurrentSpecialsExcelProcessed, 60000);
setInterval(mainEventLoop, 60000);
displaySchedule();

btnDirOneDriveToProcess.addEventListener('click', function(){
  dirOneDriveToProcess();
})

btnTidyUpProcessed.addEventListener('click', async function(){
  await tidyUpProcessed();
})

btnTidyUpOriginals.addEventListener('click', async function(){
  await tidyUpOriginals();
})

btnRunFirstExcel.addEventListener('click', async function(){
  doRunning = true;
  await runFirstExcel();
})

btnFileProcessed.addEventListener('click', async function (){
  await isCurrentExcelProcessed();
})

btnCopyOriginals.addEventListener('click', async function(){
  await copyOriginalsToOneDrive();
})

btnCopyProcessed.addEventListener('click', async function(){
  await copyProcessedToOneDrive();
})

btnCopySpecialsSheets.addEventListener('click', async function(){
  await copySpecialsSheets();
})

btnTest.addEventListener('click', async function(){
  //doRunning = false;
  //let mySpecials = await readSpecialProgress();
  //await saveSpecialProgress(mySpecials);
  //let stats = await specialFileProgressStats()
  //await moveOldSpecialsToProcessed();
  //await tidyProcessedSpecials();
  //await checkStatus();
  //let myLookahead = await lookahead();
  //console.log(myLookahead);
  //await calculateSpecialProgressV2();
  //await specialProgressCompleteV2();
  //console.log(await copyOneDriveToProcessToDropbox());
  //await findMessageInLog(msgCopyToProcessOneDriveToDropbox);
  //console.log(await getEventDetails(eventStartProcessingSpecialSheets));
  await copyMastersToToProcessOneDrive();
 })

btnDeleteOldFolders.addEventListener('click', async function(){
  await deleteOldFolders();
})

btnDeleteToProcessOneDrive.addEventListener('click', async function(){
  await deleteOneDriveToProcess();
})

btnRunSpecialExcel.addEventListener('click', async function(){
  doSpecials = true;
  await runSpecialsExcel();
})

btnToggleSchedule.addEventListener('click', async function(){
  if (divSchedule.style.display == "block"){
    divSchedule.style.display = "none";
  } else {
    divSchedule.style.display = "block";
    await displaySchedule();
  }
})

btnToggleButtons.addEventListener('click', async function(){
  if (divButtons.style.display == "block"){
    divButtons.style.display = "none";
  } else {
    divButtons.style.display = "block";
  }
})

async function getSettings(){
  const data = await window.fs.readFileSync('./settings.json');
  const mySettings = JSON.parse(data);
  return mySettings;
}
async function myPath(){
  const response = await window.versions.ping();
  return response;
}

async function dirOneDriveToProcess(){
  const theSettings = await getSettings();
  console.log(theSettings.toProcessOneDrive);
  const myFiles = await window.fs.readdirSyncExcel(theSettings.toProcessOneDrive);
  statusList.innerText = 'To Process Files on OneDrive to be copied to Dropbox: ' + myFiles.length + " to do."
  let logMessage = msgCopyToProcessOneDriveToDropbox;
  let logDetails = ["Files copied: " + myFiles.length]
  for(let myFile of myFiles){
    statusList.innerText += '\r\n' + myFile;
    console.log(myFile);
    await window.fs.copyFileSync(theSettings.toProcessOneDrive, myFile, theSettings.toProcessDropbox, myFile);
    logDetails.push(myFile);
  }
  await addEventToLog(logMessage, logDetails);
  return true;
}

async function checkOneDriveToProcess(doLog = true){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.toProcessOneDrive);
  if (doLog){
    let logMessage = "Checking number of 'To Process' files on OneDrive"
    let logDetails = ["Files present: " + myFiles.length]
    await addEventToLog(logMessage, logDetails);
  }
  return myFiles.length
}

async function checkDropboxToProcess(doLog = true){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.toProcessDropbox);
  if (doLog){
    let logMessage = "Checking number of 'To Process' files on Dropbox"
    let logDetails = ["Files present: " + myFiles.length]
    await addEventToLog(logMessage, logDetails);
  }
  return myFiles.length
}

async function tidyUpFolder(folderName, message){
  let yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  let yesterdayString = formatDate(yesterday);
  console.log(folderName);
  let logMessage = "Tidy up: " + message;
  let newFolder = await window.fs.createFolderSync(folderName, yesterdayString);
  const myFiles = await window.fs.readdirSyncExcel(folderName);
  statusList.innerText = message;
  let logDetails =["Files processed: " + myFiles.length + " Folder: " + yesterdayString];
  for(let myFile of myFiles){
    statusList.innerText += '\r\n' + myFile;
    console.log(myFile);
    await window.fs.moveFileSync(folderName, myFile, newFolder, myFile);
    logDetails.push(myFile);
  }
  await addEventToLog(logMessage, logDetails);
}

function formatDate(date) {
  date = new Date(date);

  var day = ('0' + date.getDate()).slice(-2);
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var year = date.getFullYear();

  return year + '-' + month + '-' + day;
}

async function runFirstExcel(){
  const theSettings = await getSettings();
  console.log(theSettings.toProcessDropbox);
  const myFiles = await window.fs.readdirSyncExcel(theSettings.toProcessDropbox);
  statusList.innerText = 'First Excel on Dropbox'
  statusList.innerText += '\r\n' + myFiles[0];
  currentExcel = myFiles[0];
  currentExcelInfo.innerText = currentExcel;
  if (currentExcel != undefined){
    let logMessage = "Updating Excel"
    let logDetails = [myFiles[0]];
    await window.fs.runFirstExcel(theSettings.toProcessDropbox, myFiles[0]);
    await addEventToLog(logMessage, logDetails);
  } else {
    let logMessage = "All Excel Updates Complete"
    let logDetails = [];
    await addEventToLog(logMessage, logDetails);
  }
  finishedCounter = 0;
}

async function isCurrentExcelProcessed(){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.processedDropbox);
  let found = false;
  console.log(finishedCounter);
  if (currentExcel != undefined){
    for(let myFile of myFiles){
      if (myFile == currentExcel){
        found = true
        statusList.innerText += '\r\n' + myFile + " ...is processed";
        currentExcel = undefined;
        currentExcelInfo.innerText = "";
        finishedCounter += 1;
        let logMessage = "Excel Update Complete";
        let logDetails =[myFile];
        await addEventToLog(logMessage, logDetails);
      }
    }
    if (!found){
      statusList.innerText += '\r\n' + currentExcel + " ...still being processed";
    }
  } else {
    finishedCounter += 1;
    if (finishedCounter >= 2){
      if (doRunning){
        await runFirstExcel();
      } else {
        statusList.innerText += '\r\n' + "Currently not processing";
        finishedCounter = 0;
      }
    }
  }
}

async function copyOriginalsToOneDrive(){
  const theSettings = await getSettings();
  console.log(theSettings.originalsDropbox);
  const myFiles = await window.fs.readdirSyncExcel(theSettings.originalsDropbox);
  statusList.innerText = 'Originals on Dropbox'
  let logMessage = " Copying Originals from Dropbox to One Drive"
  let logDetails =["Files copied: " + myFiles.length]
  for(let myFile of myFiles){
    statusList.innerText += '\r\n' + myFile;
    console.log(myFile);
    await window.fs.copyFileSync(theSettings.originalsDropbox, myFile, theSettings.originalsOneDrive, myFile);
    statusList.innerText += ' ...copied to OneDrive';
    logDetails.push(myFile);
  }
  await addEventToLog(logMessage, logDetails);
}

async function copyProcessedToOneDrive(){
  const theSettings = await getSettings();
  console.log(theSettings.processedDropbox);
  const myFiles = await window.fs.readdirSyncExcel(theSettings.processedDropbox);
  statusList.innerText = 'Processed on Dropbox'
  let logMessage = " Copying Processed from Dropbox to One Drive"
  let logDetails =["Files copied: " + myFiles.length]
  for(let myFile of myFiles){
    statusList.innerText += '\r\n' + myFile;
    console.log(myFile);
    await window.fs.copyFileSync(theSettings.processedDropbox, myFile, theSettings.processedOneDrive, myFile);
    statusList.innerText += ' ...copied to OneDrive';
    logDetails.push(myFile);
  }
  await addEventToLog(logMessage, logDetails);
}

async function copyMastersToToProcessOneDrive(){
  const theSettings = await getSettings();
  console.log(theSettings.processedOneDrive);
  const myFiles = await window.fs.readdirSyncExcel(theSettings.processedOneDrive);
  const masterFiles = myFiles.filter((myFile) => {
    return myFile.startsWith('MASTER_');
  })
  statusList.innerText = 'Master Sheets Processed on One Drive'
  let logMessage = " Copying Masters Processed To To Process One Drive"
  let logDetails =["Files copied: " + masterFiles.length]
  for(let myMaster of masterFiles){
    statusList.innerText += '\r\n' + myMaster;
    console.log(myMaster);
    await window.fs.copyFileSync(theSettings.processedOneDrive, myMaster, theSettings.toProcessOneDrive, myMaster);
    statusList.innerText += ' ...copied to OneDrive';
    logDetails.push(myMaster);
  }
  await addEventToLog(logMessage, logDetails);
}

async function copySpecialsSheets(){
  const theSettings = await getSettings();
  statusList.innerText = 'Special Files'
  let totalNum = 0;
  let copied = 0;
  let alreadyPresent = 0;
  let excelFilenames = [];
  for(let myChannel of theSettings.specialsSheets){
    const myFiles = await window.fs.readdirSyncExcel(myChannel.folder);
    statusList.innerText += '\r\n' + myChannel.channel;
    for(let myFile of myFiles){
      totalNum += 1;
      let isAlreadyCopied = await alreadyCopiedV2(myChannel.channel, myFile);
      if (!isAlreadyCopied){
        statusList.innerText += '\r\n' + myFile;
        console.log(myFile);
        await window.fs.copyFileSync(myChannel.folder, myFile, theSettings.specialsDropbox, myFile);
        statusList.innerText += ' ...copied to Specials Dropbox';
        copied += 1;
        excelFilenames.push(myFile);
      } else {
        statusList.innerText += '\r\n' + myFile + ' ...already copied';
        alreadyPresent += 1
      }
    }
  }
  await calculateSpecialProgressV2();
  let logMessage = "Copied Specials Files"
  let logDetails = []
  logDetails.push("Number of files: " + totalNum);
  logDetails.push("Number copied: " + copied);
  logDetails.push("Number already present: " + alreadyPresent);
  for(let fname of excelFilenames){
    logDetails.push("File copied: " + fname)
  }
  if (copied > 0){
    addEventToLog(logMessage, logDetails);
  }
  
  return await specialProgressCompleteV2();
}

async function getTime(){
  let myTime = await window.tc.nowAsTimecode();
  currentTime.innerText = myTime;
}
async function readSpecialProgress(){
  let filename = './specialProgress.json';
  let mySpecialProgress = await readJson(filename);
  if (mySpecialProgress == null) {
    mySpecialProgress = {}
    mySpecialProgress.date = getToday();
    mySpecialProgress.progress = []
    console.log(mySpecialProgress);
  }
  return mySpecialProgress;
}

async function saveSpecialProgress(specialProgress){
  let filename = './specialProgress.json'
  await saveJson(specialProgress, filename);
}

async function readConditions(){
  let filename = './conditions.json';
  let myConditions = await readJson(filename);
  if (myConditions == null) {
    myConditions = {}
    myConditions.date = getToday();
    myConditions.conditions = []
  } else {
    if (myConditions.date != getThirtyHourToday().toISOString()){
      myConditions.date = getThirtyHourToday();
      myConditions.conditions = []
    }
  }
  return myConditions;
}

async function saveSpecialProgress(specialProgress){
  let filename = './specialProgress.json'
  await saveJson(specialProgress, filename);
}

async function saveConditions(objConditions){
  let filename = './conditions.json'
  await saveJson(objConditions, filename);
}

async function saveJson(myObject, filename){
  if (await window.fs.existsSync(filename)){
    await window.fs.unlinkSync(filename);
  }
  let data = JSON.stringify(myObject, null, 2);
  await window.fs.writeFileSync(filename, data);
}

async function readJson(filename){
  if (await window.fs.existsSync(filename)){
    const data = await window.fs.readFileSync(filename);
    if (data != undefined){
      return JSON.parse(data);
    } else {
      return null;
    }
  } else {
    return null;
  }
}
function getToday(){
  let myDate = new Date()
  myDate.setHours(0,0,0,0);
  return myDate;
}

function getThirtyHourToday(){
  let myDate = new Date();
  let hour = myDate.getHours();
  if (hour < 6){
    myDate.setDate(myDate.getDate() - 1);
  }
  myDate.setHours(0,0,0,0);
  return myDate;
}


function alreadyCopied(channel, filename, specialProgress){
  
  for(let myProgress of specialProgress.progress){
    if (myProgress.channel == channel){
      for(let myFile of myProgress.files){
        if (myFile.file == filename){
          if ((myFile.status == 'Copied') || (myFile.status = 'Processed')){
            return true
          }
        }
      }

    }
  }
  return false
}

async function alreadyCopiedV2(channel, filename){
  let specialProgress = await readSpecialProgressV2();
  for(let myProgress of specialProgress.progress){
    if (myProgress.channel == channel){
      for(let myFile of myProgress.files){
        if (myFile.file == filename){
          if ((myFile.status == 'Copied') || (myFile.status = 'Processed')){
            return true;
          }
        }
      }

    }
  }
  return false;
}

async function deleteOneDriveToProcess(){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.toProcessOneDrive);
  statusList.innerText = 'Deleting To Process Files on OneDrive: ' + myFiles.length + " to delete"
  let logMessage = "Deleting files from OneDrive 'To Process"
  let logDetails = ["Files deleted: " + myFiles.length];
  for(let myFile of myFiles){
    statusList.innerText += '\r\n' + myFile;
    await window.fs.unlinkSyncWithFolder(theSettings.toProcessOneDrive, myFile);
    logDetails.push(myFile)
  }
  await addEventToLog(logMessage, logDetails);
}

async function runSpecialsExcel(){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.specialsDropbox);
  currentSpecialsExcel = myFiles[0];
  currentExcelInfo.innerText = currentSpecialsExcel;
  if (currentSpecialsExcel != undefined){
    statusList.innerText = 'First Specials Excel on Dropbox'
    statusList.innerText += '\r\n' + myFiles[0];
    let message = "Specials file run";
    let details = [myFiles[0]];
    await window.fs.runFirstExcel(theSettings.specialsDropbox, myFiles[0]);
    await addEventToLog(message, details);
  } else {
    doSpecials = false;
    statusList.innerText += '\r\n' + "Specials Run Complete";
    let message = "Specials file run complete";
    let details = [];
    await addEventToLog(message, details);
  }
  finishedSpecialsCounter = 0;
}

async function isCurrentSpecialsExcelProcessed(){
  const theSettings = await getSettings();
  await calculateSpecialProgressV2();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.specialsProcessedDropbox);
  let found = false;
  if (currentSpecialsExcel != undefined){
    for(let myFile of myFiles){
      if (myFile == currentSpecialsExcel){
        let message = "Specials file processed";
        let details = [myFile];
        await addEventToLog(message, details);
        found = true
        statusList.innerText += '\r\n' + myFile + " ...is processed";
        currentSpecialsExcel = undefined;
        currentExcelInfo.innerText = "";
        finishedSpecialsCounter += 1;
        console.log (new Date().toLocaleTimeString() + ": Finished Specials Counter: " + finishedCounter)
      }
    }
    if (!found){
      statusList.innerText += '\r\n' + currentSpecialsExcel + " ...still being processed";
      console.log (new Date().toLocaleTimeString() + ": Finished Specials Counter: " + finishedCounter)
    }
  } else {
    finishedSpecialsCounter += 1;
    console.log (new Date().toLocaleTimeString() + ": Finished Specials Counter: " + finishedCounter)
    if (finishedSpecialsCounter >= 2){
      if (doSpecials){
        await runSpecialsExcel();
      } else {
        statusList.innerText += '\r\n' + "Specials Currently not processing";
        finishedSpecialsCounter = 0;
      }
    }
  }
}

async function deleteFolders(path, numDays){
  const myFiles = await window.fs.readdirSyncFolders(path);
  const myToday = getToday();
  statusList.innerText += '\r\n Deleting folders on ' + path 
  statusList.innerText += '\r\n ' + myFiles.length + " to check"
  let counter = 0
  let foldersDeleted = [];
  for(let myFile of myFiles){
    const folderDate = new Date(myFile);
    let diff = daysDiff(myToday, folderDate);
    if (diff > numDays){
      await window.fs.rmSync(path, myFile);
      console.log(folderDate, diff);
      counter += 1
      foldersDeleted.push(myFile);
      statusList.innerText += '\r\n Folder: ' +  myFile + " ...deleted.";
    }
  }
  statusList.innerText += '\r\n' +  counter + " folder(s) deleted.";
  let logMessage = "Deleting old folders";
  let logDetails = ["Folders deleted: " + counter];
  logDetails.push("Path: " + path);
  for(let folderName of foldersDeleted){
    logDetails.push(folderName);
  }
  await addEventToLog(logMessage, logDetails);
}

function daysDiff(date1, date2){
  return Math.floor((date1 - date2) / 86400000);
}

async function readLogFile(){
  let filename = logFileName();
  console.log(filename);
  if (await window.fs.existsSync(filename)){
    const data = await window.fs.readFileSync(filename);
    if (data === undefined){
      let myLog = {};
      myLog.event = [];
      return myLog;
    } else {
      return JSON.parse(data);
    }
  } else {
    let myLog = {};
    myLog.event = [];
    return myLog;
  }
}

async function saveLogFile(myLog){
  let filename = logFileName();
  if (await window.fs.existsSync(filename)){
    await window.fs.unlinkSync(filename);
  }
  let data = JSON.stringify(myLog, null, 2);
  await window.fs.writeFileSync(filename, data);
  console.log(filename)
  console.log(data);
}

async function addEventToLog(message, details){
  let currentLog = await readLogFile();
  let newEvent = {};
  newEvent.time = new Date();
  newEvent.message = message;
  newEvent.details = details;
  currentLog.event.push(newEvent);
  await saveLogFile(currentLog);
}


function getlogFileDate(){
  let myDate = new Date();
  let myHours = myDate.getHours();
  if (myHours >= 0 && myHours < 6){
    myDate.setDate(myDate.getDate() - 1);
  }
  myDate.setHours(0,0,0,0);
  return myDate;
}

function logFileName(){
  let datePart = formatDate(getlogFileDate()) ;
  return "./logs/log_" + datePart + ".json";
}

async function mainEventLoop(){
  await displaySchedule();
  await displayMainStatus();
  
  statusList.innerText = 'Now: ' + new Date();
  
  let isEnabled = await todayEnabled();
  if (!isEnabled.enabled){
    statusList.innerText += "\r\n Not enabled today"
  } else {
    let numExcel = await checkDropboxToProcess(false);
    statusList.innerText += '\r\n Files remaining: ' + numExcel;
    statusList.innerText += '\r\n Current Conditions:';
    let objConditions = await readConditions();
    for (let condition of objConditions.conditions){
      statusList.innerText += '\r\n' + condition;
    }
  
    statusList.innerText += '\r\n Current Poll Times:';
    for (let poll of pollTimes){
      statusList.innerText += '\r\n' + poll.time + " - " + poll.name;
    }
  
    if (!doingMainLoop){
      doingMainLoop = true;
      const theSettings = await getSettings();
      for(let myEvent of theSettings.schedule){
        console.log(new Date().toLocaleTimeString() + ": " + myEvent.event + ": " + myEvent.name);
        console.log(new Date().toLocaleTimeString() + ": " + timeFromString(myEvent.startTime) + ": " + timeFromString(myEvent.endTime));
        let runningEvent
        if (inTimeRange(myEvent)){
          if (!await alreadyDone(myEvent)){
            if (await isStartConditionMet(myEvent)){
              if (pollTimeGood(myEvent)){
                await startTheEvent(myEvent);
              }
            }
          }
        }
      }
      doingMainLoop = false;
    }
  }
}


function inTimeRange(theEvent){
  let now = new Date();
  let startTime = timeFromString(theEvent.startTime);
  let endTime = timeFromString(theEvent.endTime);
  return (now >= startTime) && (now <= endTime);
  
}

function inCurrentEvents(id){
  if (currentEvents.length == 0){
    return null;
  } else {
    for(let event of currentEvents){
      if (event.id == id){
        return event;
      }
    }
  }
}

async function isStartConditionMet(theEvent){
  switch (theEvent.startCondition){
    case "Time Only":
      return true;
      break;
    case "readyCopyOneDriveToDropbox":
      return await checkConditions("readyCopyOneDriveToDropbox");
      break;
    case "originalsTidyiedDropbox":
      return await checkConditions("originalsTidyiedDropbox");
      break;
      case "originalsTidyiedOneDrive":
        return await checkConditions("originalsTidyiedOneDrive");
        break;
    case "processedTidyiedDropbox":
      return await checkConditions("processedTidyiedDropbox");
      break;
    case "processedTidyiedOneDrive":
      return await checkConditions("processedTidyiedOneDrive");
      break;
    case "specialSheetsCopied":
      return await checkConditions("specialSheetsCopied");
      break;
    case "processingSpecialSheetsStarted":
      return await checkConditions("processingSpecialSheetsStarted");
      break;
    case "processingSpecialSheetsFinished":
      return await checkConditions("processingSpecialSheetsFinished");
      break;
    case "excelProcessingStarted":
      return await checkConditions("excelProcessingStarted");
      break;
    case "processingExcelFinished":
      return await checkConditions("processingExcelFinished");
      break;
    case "copyiedOriginalsDropboxToOneDrive":
      return await checkConditions("copyiedOriginalsDropboxToOneDrive");
      break;
    case "copyiedProcessedDropboxToOneDrive":
      return await checkConditions("copyiedProcessedDropboxToOneDrive");
      break;
    case "toProcessOneDriveDeleted":
      return await checkConditions("toProcessOneDriveDeleted");
      break;
    default:
      return false;
      break;
  }
}

async function alreadyDone(theEvent){
  return checkConditions(theEvent.successfulMessage);
}

async function checkConditions(condition){
  let objConditions = await readConditions();
  for(let test of objConditions.conditions){
    if (test == condition){
      return true;
    }
  }
  return false;
}

async function startTheEvent(theEvent){
  switch (theEvent.functionName){
    case "checkOneDriveToProcess":
      let numFiles = await checkOneDriveToProcess();
      if (processEndCondition(theEvent.endCondition, numFiles)){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
      } else {
        await addCondition(theEvent.errorMessage, null);
      }
      break;
    case "copyOneDriveToProcessToDropbox":
      let result = await copyOneDriveToProcessToDropbox();
      if (result){
        //await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
      } else {
        await addCondition(theEvent.errorMessage, null);
      }
      break;
    case "tidyOriginalsDropbox":
      if (await tidyUpOriginals("Dropbox")){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
      } else {
        await addCondition(theEvent.errorMessage, null);
      }
      break;
    case "tidyOriginalsOneDrive":
      if (await tidyUpOriginals("OneDrive")){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
      } else {
        await addCondition(theEvent.errorMessage, null);
      }
      break;
    case "tidyProcessedDropbox":
      if (await tidyUpProcessed("Dropbox")){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
      } else {
        await addCondition(theEvent.errorMessage, null);
      }
      break;
    case "tidyProcessedOneDrive":
      if (await tidyUpProcessed("OneDrive")){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
      } else {
        await addCondition(theEvent.errorMessage, null);
      }
      break;
    case "deleteOldFolders":
      if (await deleteOldFolders()){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
      } else {
        await addCondition(theEvent.errorMessage, null);
      }
      break;
    case "copySpecialSheets":
      let stats = await copySpecialsSheets();
      let done = true;
      let doneV2 = true;
      for (let myStat of stats){
        if(myStat.lookaheadCopied == 0) {
          done = false;
        }
        if(myStat.lookaheadCopiedV2 == 0) {
          doneV2 = false;
        }
      }
      console.log("In case CopySpecialSheets");
      console.log("=========================");
      console.log("stats:", stats);
      console.log("done:", done);
      console.log("doneV2:", doneV2);
      if (!doneV2){
        let myCuttOff = await cutOffTime();
        console.log(myCuttOff);
        let now = new Date();
        console.log(now);
        if (now > myCuttOff){
          console.log("Passed cut off: " + myCuttOff.toLocaleTimeString());
          doneV2 = true;
        } else {
          console.log("Not passed cut off: " + myCuttOff.toLocaleTimeString());
        }
      }
      if (doneV2){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
        removeEventFromPollTimes(theEvent);
      } else {
        await addCondition(theEvent.errorMessage, null);
        addEventToPollTimes(theEvent);
      }
      break;
    case eventStartProcessingSpecialSheets:
      doSpecials = true;
      await runSpecialsExcel();
      await addCondition(theEvent.successfulMessage, null);
      break;
    case "isProcessingSpecialSheetsFinished":
      let myStats = await specialProgressCompleteV2();
      let isDone = true;
      for (let myStat of myStats){
        if (myStat.statusV2 != "Ready"){
          let myCuttOff = await cutOffTime();
          let now = new Date();
          if (now > myCuttOff){
            if (myStat.copied == 0){
              isDone = true
            }
          } else {
            isDone = false;
          }
        }
      }
      if (isDone){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
        removeEventFromPollTimes(theEvent);
      } else {
        await addCondition(theEvent.errorMessage, null);
        addEventToPollTimes(theEvent);
      }
      break;
    case eventStartExcelProcessing:
      doRunning = true;
      await runFirstExcel();
      await addCondition(theEvent.successfulMessage, null);
      break;
    case "isProcessingExcelFinished":
      let numExcel = await checkDropboxToProcess();
      if (processEndCondition(theEvent.endCondition, numExcel)){
        await addCondition(theEvent.successfulMessage, theEvent.errorMessage);
        doRunning = false;
        removeEventFromPollTimes(theEvent);
      } else {
        await addCondition(theEvent.errorMessage, null);
        addEventToPollTimes(theEvent);
      }
      break;
    case "copyOriginalsDropboxToOneDrive":
      await copyOriginalsToOneDrive();
      await addCondition(theEvent.successfulMessage, null);
      break;
    case "copyProcessedDropboxToOneDrive":
      await copyProcessedToOneDrive();
      await addCondition(theEvent.successfulMessage, null);
      break;
    case "deleteToProcessOneDrive":
      await deleteOneDriveToProcess();
      await addCondition(theEvent.successfulMessage, null);
      break;
    case "copyMastersToToProcessOneDrive":
      await copyMastersToToProcessOneDrive();
      await addCondition(theEvent.successfulMessage, null);
      break;
  }
  
}

async function addCondition(toAdd, toRemove){
  let objConditions = await readConditions();
  let conditions = objConditions.conditions
  if (toAdd != null){
    conditions = conditions.filter(conditions => conditions !== toAdd);
    conditions.push(toAdd);
  }
  if (toRemove != null){
    conditions = conditions.filter(conditions => conditions !== toRemove);
  }
  objConditions.conditions = conditions;
  await saveConditions(objConditions);
}

function processEndCondition(endCondition, value){
  if(endCondition.comparison == "greaterThan"){
    return value > Number(endCondition.value)
  }
  if(endCondition.comparison == "equal"){
    return value == Number(endCondition.value)
  }
 
}
function timeFromString(timeString){
  const [hour, minutes, seconds] = timeString.split(':');
  const theStart = getThirtyHourToday();
  theStart.setHours(hour);
  theStart.setMinutes(minutes);
  theStart.setSeconds(seconds);
  return theStart;
}

async function tidyUpProcessed(theFolderName){
  const theSettings = await getSettings();
  if (theFolderName == "Dropbox"){
    await tidyUpFolder(theSettings.processedDropbox, 'Processed Files on Dropbox'); 
  } else {
    await tidyUpFolder(theSettings.processedOneDrive, 'Processed Files on OneDrive');
  }
  return true;
}

async function tidyUpOriginals(theFolderName){
  const theSettings = await getSettings();
  if (theFolderName == "Dropbox"){
    await tidyUpFolder(theSettings.originalsDropbox, 'Original Files on Dropbox')
  } else {
    await tidyUpFolder(theSettings.originalsOneDrive, 'Original Files on One Drive')
  }
  return true;
}

async function deleteOldFolders(){
  const theSettings = await getSettings();
  await deleteFolders(theSettings.processedDropbox, 14)
  await deleteFolders(theSettings.processedOneDrive, 14);
  await deleteFolders(theSettings.originalsDropbox, 14)
  await deleteFolders(theSettings.originalsOneDrive, 14);
  await deleteFolders(theSettings.specialsProcessedDropbox, 14);
  return true;
}

async function specialsFileProcessed(filename){
  const theSettings = await getSettings();
  let specialProgress = await readSpecialProgress();
  if (specialProgress.date == getThirtyHourToday().toISOString()){
    let channelData = theSettings.specialsSheets
    for (let myChannel of channelData){
      if (filename.startsWith(myChannel.prefix)){
        if(setSpecialFileInProgressProcessed(myChannel.channel, filename, specialProgress)){
          await saveSpecialProgress(specialProgress);
          return true;
        }
      }
    }
  }
}

function setSpecialFileInProgressProcessed(channel, filename, specialProgress){
  
  for(let myProgress of specialProgress.progress){
    if (myProgress.channel == channel){
      for(let myFile of myProgress.files){
        if (myFile.file == filename){
          if (myFile.status == 'Copied'){
            myFile.status = "Processed"
            return true
          }
        }
      }
    }
  }
  return false
}

async function specialFileProgressStats(){
  const theSettings = await getSettings();
  let stats = [];
  let channelData = theSettings.specialsSheets
  for (let myChannel of channelData){
    let thisChannel = {};
    thisChannel.channel = myChannel.channel;
    thisChannel.processed = 0
    thisChannel.copied = 0
    stats.push(thisChannel);
  }
  let specialProgress = await readSpecialProgress();
  if (specialProgress.date == getThirtyHourToday().toISOString()){
    for(let myProgress of specialProgress.progress){
      for(let myFile of myProgress.files){
        for (let myChannel of stats){
          if (myChannel.channel == myProgress.channel){
            if (myFile.status == "Copied"){
              myChannel.copied += 1
            }
            if (myFile.status == "Processed"){
              myChannel.processed += 1
            }
          }
        }
      }
    }
  }
  return stats
}

function addMinutes(date, minutes) {
  return new Date(date.getTime() + minutes*60000);
}

function addEventToPollTimes(theEvent){
  let newPollTime = addMinutes(new Date(), Number(theEvent.pollTimeMinutes));
  pollTimes = pollTimes.filter(pollTimes => pollTimes.name !== theEvent.functionName);
  let newPoll = {};
  newPoll.name = theEvent.functionName;
  newPoll.time = newPollTime;
  pollTimes.push(newPoll);
}

function removeEventFromPollTimes(theEvent){
  pollTimes = pollTimes.filter(pollTimes => pollTimes.name !== theEvent.functionName);
}

function pollTimeGood(theEvent){
  for (let myTime of pollTimes){
    if (myTime.name == theEvent.functionName){
      let now = new Date();
      return (now > myTime.time);
    }
  }
  return true;
}


function addFileToSpecialProgress(specialProgress, channel, file, status){
  let foundFile = false;
  let foundChannel = false;
  for (let myProgress of specialProgress.progress){
    if(myProgress.channel == channel){
      foundChannel = true;
      for (let myFile of myProgress.files){
        if (myFile.file == file){
          foundFile = true;
          myFile.status = status;
        }
      }
      if (!foundFile){
        let newFile = {}
        newFile.file = file;
        newFile.status = status;
        myProgress.files.push(newFile);
      }
    }
  }
  if (!foundChannel){
    let newChannel = {}
    newChannel.channel = channel;
    newChannel.files = [];
    let newFile = {}
    newFile.file = file;
    newFile.status = status;
    newChannel.files.push(newFile);
    specialProgress.progress.push(newChannel);
  }

}

async function displaySchedule(){
  let theSettings = await getSettings();
  let theSchedule = sortedSchedule(theSettings.schedule);
  scheduleColumn01.innerText = "";
  scheduleColumn02.innerText = "";
  scheduleColumn03.innerText = "";
  for(let myEvent of theSchedule){
    scheduleColumn01.innerText += myEvent.startTime + '\r\n'; 
    scheduleColumn02.innerText += myEvent.endTime + '\r\n';
    scheduleColumn03.innerText += myEvent.name + '\r\n';
  }
}

async function checkStatus(){
  let myStatus = {}
  myStatus.date = new Date();
  myStatus.todayEnabled = await todayEnabled()
  myStatus.oneDriveToProcess = await checkOneDriveToProcess(false);
  myStatus.dropboxToProcess = await checkDropboxToProcess(false);
  myStatus.copyDetails = await findMessageInLog(msgCopyToProcessOneDriveToDropbox)
  myStatus.specialProgress = await specialProgressCompleteV2();
  myStatus.sheetProgress = await getTodaysSheets();
  writeStatus(myStatus);
  return myStatus;
}

async function writeStatus(myStatus){
  await saveStatusFile(myStatus);
}

async function saveStatusFile(myStatus){
  let filename = "./status.json";
  if (await window.fs.existsSync(filename)){
    await window.fs.unlinkSync(filename);
  }
  let data = JSON.stringify(myStatus, null, 2);
  await window.fs.writeFileSync(filename, data);
  console.log(filename)
  console.log(data);
}

async function getTodaysSheets(){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.toProcessOneDrive);
  if (myFiles.length > 0){
    todaysSheets = [];
    for (let myFile of myFiles){
      let mySheet = {}
      mySheet.name = myFile;
      mySheet.readyForProcess = false;
      mySheet.processed = false;
      todaysSheets.push(mySheet);
    }
  }
  for (let mySheet of todaysSheets){
    mySheet.processed = await isSheetProcessed(mySheet.name);
    if (mySheet.processed){
      mySheet.readyForProcess = true;
    } else {
      mySheet.readyForProcess = await isSheetReady(mySheet.name);
    }
  }

  return todaysSheets;
}

async function isSheetProcessed(sheetName){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.processedDropbox);
  for (let myFile of myFiles){
    if (myFile == sheetName){
      return true;
    }
  }
  return false;
}

async function isSheetReady(sheetName){
  const theSettings = await getSettings();
  const myFiles = await window.fs.readdirSyncExcel(theSettings.toProcessDropbox);
  for (let myFile of myFiles){
    if (myFile == sheetName){
      return true;
    }
  }
  return false;
}


async function moveOldSpecialsToProcessed(){
  const theSettings = await getSettings();
  let oldFolder = theSettings.specialsOldDropbox;
  let processedFolder = theSettings.specialsProcessedDropbox;
  const myFiles = await window.fs.readdirSyncExcel(oldFolder);
  for (let myFile of myFiles){
    await window.fs.moveFileSync(oldFolder, myFile, processedFolder, myFile);
  }
 
}

async function spreadsheetDate(filename){
  const theSettings = await getSettings();
  let channelData = theSettings.specialsSheets
  for (let myChannel of channelData){
    if (filename.startsWith(myChannel.prefix)){
      let datePart = filename.replace(myChannel.prefix, "");
      let dateString = "20" + datePart.substr(0,2) + "-" + datePart.substr(2,2) + "-" + datePart.substr(4,2);
      let theDate = new Date(dateString);
      theDate.setHours(0,0,0,0);
      return theDate
    }
  }
  return null;
}

async function tidyProcessedSpecials(){
  const theSettings = await getSettings();
  let processedFolder = theSettings.specialsProcessedDropbox;
  const myFiles = await window.fs.readdirSyncExcel(processedFolder);
  let myToday = getThirtyHourToday();
  for (let myFile of myFiles){
    let myDate = await spreadsheetDate(myFile);
    console.log(myFile + ": " + myDate);
    if (myDate < myToday){
      console.log("Ready to archive");
      await tidyProcessSpecial(myDate, myFile);
    }
  }
}

async function tidyProcessSpecial(theDate, filename){
  const theSettings = await getSettings();
  let processedFolder = theSettings.specialsProcessedDropbox;
  let dateString = formatDate(theDate);
  let newFolder = await window.fs.createFolderSync(processedFolder, dateString);

  let logMessage = "Tidy up special to: " + newFolder;
  let logDetails = [];
  console.log(filename);
  await window.fs.moveFileSync(processedFolder, filename, newFolder, filename);
  logDetails.push(filename);
  
  await addEventToLog(logMessage, logDetails);
}

async function lookahead(){
  let myToday = await getTodaySpecials();
  if (myToday === null){
    return null;
  } else {
    return myToday.lookahead;
  }
}

async function lookaheadV2(channel){
  let myToday = await getTodaySpecials();
  if (myToday === null){
    return null;
  } else {
    let myLookaheads = myToday.channelLookahead;
    for (let myLookahead of myLookaheads){
      if (myLookahead.channel == channel){
        return myLookahead.lookahead;
      }
    }
    return null;
  }
}

async function cutOffTime(){
  let myToday = await getTodaySpecials();
  if (myToday === null){
    return null;
  } else {
    return timeFromString(myToday.cutOffTime);
  }
}

async function getTodaySpecials(){
  const theSettings = await getSettings();
  const weekday = ["Sunday", "Monday", "Tuesday","Wednesday", "Thursday", "Friday", "Saturday"];
  const d = new Date();
  let day = weekday[d.getDay()];
  let theDays = theSettings.specialsLookAhead;
  for (let theDay of theDays){
    if (theDay.day == day){
      return theDay;
    }
  }
  return null;
}

async function todayEnabled(){
  const theSettings = await getSettings();
  const weekday = ["Sunday", "Monday", "Tuesday","Wednesday", "Thursday", "Friday", "Saturday"];
  const d = new Date();
  let day = weekday[d.getDay()];
  let theDays = theSettings.daysEnabled;
  for (let theDay of theDays){
    if (theDay.day == day){
      return theDay;
    }
  }
  return null;
}

async function calculateSpecialProgressV2(){
  await moveOldSpecialsToProcessed();
  await tidyProcessedSpecials();
  const theSettings = await getSettings();
  let specialsToDoFolder = theSettings.specialsDropbox;
  let processedFolder = theSettings.specialsProcessedDropbox;
  const toDoFiles = await window.fs.readdirSyncExcel(specialsToDoFolder);
  const processedFiles = await window.fs.readdirSyncExcel(processedFolder);
  let specialProgress = {};
  specialProgress.date = getThirtyHourToday();
  specialProgress.progress = [];
  let channelData = theSettings.specialsSheets
  for (let myChannel of channelData){
    let myProgress = {}
    myProgress.channel = myChannel.channel;
    myProgress.files = [];
    for (let toDoFile of toDoFiles){
      if (toDoFile.startsWith(myChannel.prefix)){
        let myItem = {};
        myItem.file = toDoFile;
        myItem.status = "Copied"
        myProgress.files.push(myItem);
      }
    }
    for (let processedFile of processedFiles){
      if (processedFile.startsWith(myChannel.prefix)){
        let myItem = {};
        myItem.file = processedFile;
        myItem.status = "Processed"
        myProgress.files.push(myItem);
      }
    }
    specialProgress.progress.push(myProgress);
  }
  let filename = './specialProgressV2.json'
  await saveJson(specialProgress, filename);
}
async function readSpecialProgressV2(){
  let filename = './specialProgressV2.json';
  let mySpecialProgress = await readJson(filename);
  if (mySpecialProgress == null) {
    mySpecialProgress = {}
    mySpecialProgress.date = getToday();
    mySpecialProgress.progress = []
    console.log(mySpecialProgress);
  }
  return mySpecialProgress;
}

async function specialProgressCompleteV2(){
  let specialProgress = await readSpecialProgressV2();
  let lookaheadDate = getThirtyHourToday();
  lookaheadDate.setDate(lookaheadDate.getDate() + await lookahead());
  console.log("Lookahead date: " + lookaheadDate);
  let theProgress = specialProgress.progress
  let details = []
  for (let channelData of theProgress){
    let thisData = {}
    thisData.channel = channelData.channel;
    thisData.lookaheadDate = lookaheadDate;
    let lookaheadDateV2 = getThirtyHourToday();
    lookaheadDateV2.setDate(lookaheadDateV2.getDate() + await lookaheadV2(thisData.channel));
    thisData.lookaheadDateV2 = lookaheadDateV2;

    console.log('this data in specialProgressCompleteV2');
    console.log('======================================');
    console.log(thisData);

    let theFiles = channelData.files
    thisData.copied = 0;
    thisData.copiedV2 = 0;
    thisData.processed = 0;
    thisData.processedV2 = 0;
    thisData.lookaheadCopied = 0;
    thisData.lookaheadCopiedV2 = 0;
    for (let theFile of theFiles){
      let testDate = await spreadsheetDate(theFile.file)
      console.log(testDate.getTime() - lookaheadDate.getTime());
      if (testDate.getTime() == lookaheadDate.getTime()){
        if (theFile.status == 'Copied'){
          thisData.lookaheadCopied += 1;
          thisData.copied += 1
        }
        if (theFile.status == 'Processed'){
          thisData.lookaheadCopied += 1;
          thisData.processed += 1
        }
      } else if (!(testDate.getTime() > lookaheadDate.getTime())){
        if (theFile.status == 'Copied'){
          thisData.copied += 1
        }
        if (theFile.status == 'Processed'){
          thisData.processed += 1
        }
      }
    

      //V2
      if (testDate.getTime() == lookaheadDateV2.getTime()){
        if (theFile.status == 'Copied'){
          thisData.lookaheadCopiedV2 += 1;
          thisData.copiedV2 += 1
        }
        if (theFile.status == 'Processed'){
          thisData.lookaheadCopiedV2 += 1;
          thisData.processedV2 += 1
        }
      } else if (!(testDate.getTime() > lookaheadDateV2.getTime())){
        if (theFile.status == 'Copied'){
          thisData.copiedV2 += 1
        }
        if (theFile.status == 'Processed'){
          thisData.processedV2 += 1
        }
      }
    }


    if ((thisData.processed > 0) && (thisData.copied == 0) && (thisData.lookaheadCopied > 0)){
      thisData.status = "Ready"
    } else {
      thisData.status = "Not Ready"
    }

    if ((thisData.processedV2 > 0) && (thisData.copiedV2 == 0) && (thisData.lookaheadCopiedV2 > 0)){
      thisData.statusV2 = "Ready"
    } else {
      thisData.statusV2 = "Not Ready"
    }

    details.push(thisData);
  }
  console.log(details);
  return details;
}

function sortedSchedule(schedule){
  let sortedData;
  sortedData = schedule.sort(function(a,b){
    return timeFromString(a.startTime) - timeFromString(b.startTime);
  });
  return sortedData;
}

function countLookaheadsCopied(theSpecialProgress){
  let myCount = 0;
  for (let theChannel of theSpecialProgress){
    myCount += theChannel.lookaheadCopied
  }

  return myCount;
}

function specialsStats(theSpecialProgress){
  let numLookahead = 0;
  let numCopied = 0;
  let numReady = 0;
  let numNotReady = 0;
  let lookaheadDate;
  let totalLookahead = theSpecialProgress.length;
  for (let theChannel of theSpecialProgress){
    lookaheadDate = new Date(theChannel.lookaheadDate).toLocaleDateString();;
    numLookahead += theChannel.lookaheadCopied;
    numCopied += theChannel.copied;
    if (theChannel.statusV2 == "Ready"){
      numReady += 1;
    } else {
      numNotReady += 1;
    }
  }
  return numCopied + " copied.\r\n Lookahead (" + lookaheadDate + "): " + numLookahead + " copied of " + totalLookahead + ".\r\n Channels: " + numReady + " processed, " + numNotReady + " not ready";
}

async function toProcessFilesToCopy(){
  const theSettings = await getSettings();
  const oneDriveFiles = await window.fs.readdirSyncExcel(theSettings.toProcessOneDrive);
  const dropboxFiles = await window.fs.readdirSyncExcel(theSettings.toProcessDropbox);
  let filesToCopy = [];
  for (let oneDriveFile of oneDriveFiles){
    let found = false;
    let isModified = false;
    for (let dropboxFile of dropboxFiles){
      if (oneDriveFile == dropboxFile){
        let oneDriveModified = await window.fs.fileModifiedTime(theSettings.toProcessOneDrive, oneDriveFile);
        let dropboxModified = await window.fs.fileModifiedTime(theSettings.toProcessDropbox, dropboxFile);
        if (oneDriveModified == dropboxModified){
          found = true;
        } else {
          isModified =true
        }
      }
    }
    if (!found){
      let fileDetails = {}
      fileDetails.filename = oneDriveFile;
      fileDetails.isModified = isModified
      filesToCopy.push(fileDetails);
    }
  }
  return filesToCopy;
}

async function copyOneDriveToProcessToDropbox(){
  const theSettings = await getSettings();
  let fileDetails = await toProcessFilesToCopy();
  statusList.innerText += '\r\n To Process Files on OneDrive to be copied to Dropbox: ' + fileDetails.length + " to do."
  let logMessage = msgCopyToProcessOneDriveToDropbox;
  let logDetails = ["Files copied: " + fileDetails.length];
  for(let myFile of fileDetails){
    let fileText = myFile.filename;
    if (myFile.isModified){
      fileText += ". File has been modified"
      await window.fs.unlinkSyncWithFolder(theSettings.toProcessDropbox, myFile.filename);
    } else {
      fileText += ". New file"
    }
    statusList.innerText += '\r\n' + fileText;
    
    await window.fs.copyFileSync(theSettings.toProcessOneDrive, myFile.filename, theSettings.toProcessDropbox, myFile.filename);
    logDetails.push(myFile);
  }
  if (fileDetails.length > 0){
    await addEventToLog(logMessage, logDetails);
  }
  return true;
}

async function findMessageInLog(theMessage){
  const theLog = await readLogFile();
  let results = theLog.event.filter(result => result.message == theMessage);
  console.log(results);
  return results;
}

async function displayMainStatus(){
  const mainStatus = await checkStatus();
  const startSpecialsEventDetails = await getEventDetails(eventStartProcessingSpecialSheets);
  const startExcelProcessingDetails = await getEventDetails(eventStartExcelProcessing);
  divMainStatus.innerText = "Status at: " + mainStatus.date.toLocaleTimeString();
  divMainStatus.innerText += "\r\n " + mainStatus.todayEnabled.day;
  if (!mainStatus.todayEnabled.enabled){
    divMainStatus.innerText += ": Not enabled"
  } else {
    divMainStatus.innerText += ": Enabled"
    divMainStatus.innerText += "\r\n To process files on OneDrive: " + mainStatus.oneDriveToProcess;
    divMainStatus.innerText += "\r\n To process files on Dropbox: " + mainStatus.dropboxToProcess;
    for (let detail of mainStatus.copyDetails){
      divMainStatus.innerText += "\r\n" + detail.details[0] + " at " + new Date(detail.time).toLocaleTimeString();
    }
    divMainStatus.innerText += "\r\n Number of Special Sheets: " + specialsStats(mainStatus.specialProgress);

    if (new Date().getTime() < startSpecialsEventDetails.startTime.getTime()){
      divMainStatus.innerText += "\r\n Specials Processing will start at: " + startSpecialsEventDetails.startTime.toLocaleTimeString() + " if all lookahead sheets copied."
    } else {
      if (!startSpecialsEventDetails.startConditionMet && startSpecialsEventDetails.inTimeRange){
        let myCuttOff = await cutOffTime();
        console.log("cutOfff")
        console.log(myCuttOff);
        divMainStatus.innerText += "\r\n Specials Processing will start once all lookahead sheets copied, or if cut-off time is passed: " + myCuttOff.toLocaleTimeString();
      }
    }

    if (new Date().getTime() < startExcelProcessingDetails.startTime.getTime()){
      divMainStatus.innerText += "\r\n Excel Processing will start at: " + startExcelProcessingDetails.startTime.toLocaleTimeString() + " if all special sheets processed"
    } else {
      if (!startExcelProcessingDetails.startConditionMet && startExcelProcessingDetails.inTimeRange){
        divMainStatus.innerText += "\r\n Excel Processing will start once all special sheets processed."
      }
    }
  }
}


async function getEventDetails(eventFunctionName){
  const theSettings = await getSettings();
  let details = {}
  for(let myEvent of theSettings.schedule){
    if (myEvent.functionName == eventFunctionName){
      details.name = myEvent.name;
      details.startTime = timeFromString(myEvent.startTime);
      details.inTimeRange = inTimeRange(myEvent);
      details.alreadyDone = await alreadyDone(myEvent);
      details.startConditionMet = await isStartConditionMet(myEvent);
    }
  }
  return details
}