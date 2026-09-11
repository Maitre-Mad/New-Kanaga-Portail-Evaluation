const SPREADSHEET_ID = '1WjsYi4GQsYyHa-LlCzg6FOr4qv9EPM7X8f6vM8aLSmQ';

function doGet(e) {

  try {
    getEvaluationConfig();
  } catch(err) {
    Logger.log("Init EvaluationQuestions error: " + err.message);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Portail Kanaga')
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/forms_48dp.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gestion de Session
 */
function generateSessionToken() {
  return Utilities.getUuid();
}

function cleanupExpiredSessions() {
  const props = PropertiesService.getScriptProperties();
  const allKeys = props.getKeys();
  const now = new Date().getTime();
  allKeys.forEach(key => {
    if (key.startsWith('SESSION_')) {
      try {
        const sessionData = JSON.parse(props.getProperty(key));
        if (sessionData.expiry < now) {
          props.deleteProperty(key);
        }
      } catch (e) {
        props.deleteProperty(key);
      }
    }
  });
}

function verifySession(token) {
  if (!token) throw new Error("Accès refusé: Aucun jeton de session fourni.");
  const props = PropertiesService.getScriptProperties();
  const sessionStr = props.getProperty('SESSION_' + token);
  if (!sessionStr) throw new Error("Accès refusé: Session invalide ou expirée. Veuillez vous reconnecter.");
  
  try {
    const sessionData = JSON.parse(sessionStr);
    if (sessionData.expiry < new Date().getTime()) {
      props.deleteProperty('SESSION_' + token);
      throw new Error("Accès refusé: Session expirée. Veuillez vous reconnecter.");
    }
    return sessionData.user;
  } catch (e) {
    throw new Error("Accès refusé: Session corrompue.");
  }
}

function logout(token) {
  if (token) {
    PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
  }
  return { success: true };
}

/**
 * Authentifie un utilisateur avec son nom d'utilisateur et mot de passe.
 */
function authenticateUser(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Utilisateurs');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === username && data[i][1] === password) {
          const societiesString = data[i][3] || '';
          const societies = societiesString.split(',').map(s => s.trim()).filter(String);
          let role = data[i][5] || 'Utilisateur';
          if (data[i][0] === 'mhdicko@kanagaconsulting.com' && (!data[i][5] || data[i][5] === '')) {
            role = 'Admin';
            sheet.getRange(i + 1, 6).setValue('Admin');
          }
          const allowedProjectsStr = data[i][6] || '';
          const allowedProjectsList = allowedProjectsStr.split(',').map(p => p.trim()).filter(String);
          const userObj = {
            username: data[i][0],
            fullName: data[i][2] || data[i][0],
            societies: societies.length > 0 ? societies : ['Kanaga Consulting SARL (Mali)'],
            activeSociety: societies.length > 0 ? societies[0] : 'Kanaga Consulting SARL (Mali)',
            language: 'fr',
            role: role,
            allowedProjects: allowedProjectsList
          };

          cleanupExpiredSessions();
          const token = generateSessionToken();
          const expiry = new Date().getTime() + (12 * 60 * 60 * 1000); // 12 hours
          PropertiesService.getScriptProperties().setProperty('SESSION_' + token, JSON.stringify({
            user: userObj,
            expiry: expiry
          }));

          return {
            success: true,
            token: token,
            user: userObj
          };
        }
      }
    }
  } catch (e) {
    console.error("Erreur d'authentification : " + e);
  }
  return { success: false, message: "Identifiants incorrects" };
}

function getDropdownData(token, activeSociety) {
  const sessionUser = verifySession(token);
  const role = sessionUser.role;
  const consultantAllowedProjects = sessionUser.allowedProjects;

  try {
    if (!activeSociety) activeSociety = 'Kanaga Consulting SARL (Mali)';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const projectTasksSheet = ss.getSheetByName("ProjectTasks");
    const projectsSheet = ss.getSheetByName("Projects Odoo");
    const allowedProjects = new Set();

    if (projectsSheet) {
      const projectsData = projectsSheet.getRange(2, 2, projectsSheet.getLastRow() - 1, 3).getValues();
      projectsData.forEach(row => {
        const project = row[0];
        const society = row[2];
        
        let hasAccess = false;
        if (role === 'ConsultantExterne') {
          if (project && consultantAllowedProjects && consultantAllowedProjects.includes(project)) {
            hasAccess = true;
          }
        } else {
          if (project && (society === activeSociety || society === '')) {
            hasAccess = true;
          }
        }
        
        if (hasAccess) {
          allowedProjects.add(project);
        }
      });
    }

    const taskMap = {};
    if (projectTasksSheet) {
        const projectTasksData = projectTasksSheet.getRange(2, 1, projectTasksSheet.getLastRow() - 1, 8).getValues();
        projectTasksData.forEach(row => {
          const project = row[0];
          const task = row[2];   
          const society = row[4];
          const isCompleted = (row[5] === true || row[5] === 'true' || row[5] === 'TRUE');
          const id = row[6];
          const parentId = row[7];
          
          if (project && task && (society === activeSociety || society === '')) {
            if (!taskMap[project]) taskMap[project] = [];
            taskMap[project].push({ id: id, name: task, parentId: parentId, isCompleted: isCompleted });
          }
        });
    }
    
    const finalProjectData = {};
    allowedProjects.forEach(project => {
        finalProjectData[project] = taskMap[project] || [];
    });
    return finalProjectData;
  } catch (e) {
    console.error("Error in getDropdownData: " + e.message);
    return {};
  }
}

function getMonday(d) {
  d = new Date(d);
  var day = d.getDay(), diff = d.getDate() - day + (day == 0 ? -6 : 1);
  var monday = new Date(d.setDate(diff));
  var tzoffset = (monday.getTimezoneOffset() * 60000); 
  return new Date(monday.getTime() - tzoffset).toISOString().split('T')[0];
}

function isDurationValid(userEmail, entryDate, newDuration, editingRowId = null) {
  let totalDuration = 0;
  const entryDateString = entryDate.toLocaleDateString();
  const stagingSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Staging Data");
  if (stagingSheet) {
    const data = stagingSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userEmail && new Date(data[i][3]).toLocaleDateString() === entryDateString) {
        if (editingRowId && (i + 1) == editingRowId) continue;
        totalDuration += parseFloat(data[i][7]);
      }
    }
  }
  return (totalDuration + parseFloat(newDuration)) <= 15;
}

function submitTimesheet(token, formData) {
  const sessionUser = verifySession(token);
  const userObj = sessionUser;
  try {
    const entryDate = new Date(formData.entryDate);
    const today = new Date();
    const oneMonthAgo = new Date();
    oneMonthAgo.setMonth(today.getMonth() - 1);
    
    entryDate.setMinutes(entryDate.getMinutes() + entryDate.getTimezoneOffset());
    today.setHours(0, 0, 0, 0);
    oneMonthAgo.setHours(0, 0, 0, 0);
    
    if (entryDate > today) throw new Error("La date ne peut pas être dans le futur.");
    if (entryDate < oneMonthAgo) throw new Error("La date ne peut pas être antérieure à un mois.");
    if (parseFloat(formData.duration) <= 0) {
      throw new Error("La durée doit être supérieure à 0 heure.");
    }
    
    if (!isDurationValid(userObj.username, entryDate, formData.duration)) {
      throw new Error("La durée totale pour cette journée ne peut pas dépasser 15 heures.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rowData = [
      new Date(),
      userObj.username,
      userObj.fullName,
      entryDate,
      formData.project,
      formData.task || '',
      formData.description || '',
      parseFloat(formData.duration),
      userObj.activeSociety
    ];

    const stagingSheet = ss.getSheetByName('Staging Data');
    if (stagingSheet) {
      stagingSheet.appendRow(rowData);
    }
    
    const mondayStr = getMonday(formData.entryDate);
    const sheetName = 'Semaine du ' + mondayStr;
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        'Timestamp', 'Employee Email', 'Full Name', 'Date of Work', 
        'Projet', 'Task', 'Description', 'Temps passé (en heures)', 'Société'
      ]);
      sheet.setFrozenRows(1);
      sheet.getRange("A1:I1").setFontWeight("bold");
    }
    sheet.appendRow(rowData);
    
    return true;
  } catch (e) {
    throw new Error("Erreur serveur : " + e.message);
  }
}

function getUserTimesheets(token) {
  const sessionUser = verifySession(token);
  const username = sessionUser.username;

  const timesheets = [];
  try {
    const today = new Date();
    const mondayStr = getMonday(today);
    const sheetName = 'Semaine du ' + mondayStr;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === username) {
        timesheets.push({
          rowId: i + 1,
          date: row[3] instanceof Date ? row[3].toISOString() : row[3],
          project: row[4],
          task: row[5],
          description: row[6],
          duration: row[7]
        });
      }
    }
  } catch (e) {
    console.error("Erreur getUserTimesheets: " + e);
  }
  return timesheets;
}

function archiveTimesheets() {
  try {
    const rootFolderId = '1LDR2du1mTD6vOyn-9nxGL4ZhDgCVpCgA';
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    
    // Pattern to match "Semaine du YYYY-MM-DD"
    const regex = /^Semaine du (\d{4})-(\d{2})-(\d{2})$/;
    
    // We shouldn't archive the CURRENT week.
    const today = new Date();
    const currentMondayStr = getMonday(today);
    const currentWeekName = 'Semaine du ' + currentMondayStr;
    
    let archivedCount = 0;
    
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const name = sheet.getName();
      
      if (name === currentWeekName) continue; // Skip current week
      
      const match = name.match(regex);
      if (match) {
        const yyyy = match[1];
        const mm = parseInt(match[2], 10);
        const dd = match[3];
        
        let quarter = 1;
        if (mm >= 4 && mm <= 6) quarter = 2;
        else if (mm >= 7 && mm <= 9) quarter = 3;
        else if (mm >= 10 && mm <= 12) quarter = 4;
        
        const quarterFolderName = `Trimestre ${quarter} ${yyyy}`;
        
        // Find or create quarter folder
        let quarterFolder;
        const folders = rootFolder.searchFolders(`title = '${quarterFolderName}' and trashed = false`);
        if (folders.hasNext()) {
          quarterFolder = folders.next();
        } else {
          quarterFolder = rootFolder.createFolder(quarterFolderName);
        }
        
        // Create new Spreadsheet for the week
        const newSpreadsheet = SpreadsheetApp.create(name);
        const newFile = DriveApp.getFileById(newSpreadsheet.getId());
        newFile.moveTo(quarterFolder); // Move to quarter folder
        
        // Copy data to new spreadsheet
        const newSheet = newSpreadsheet.getSheets()[0];
        newSheet.setName(name);
        const dataRange = sheet.getDataRange();
        const data = dataRange.getValues();
        
        if (data.length > 0) {
          const targetRange = newSheet.getRange(1, 1, data.length, data[0].length);
          targetRange.setValues(data);
          try { targetRange.setBackgrounds(dataRange.getBackgrounds()); } catch(e){}
          try { targetRange.setFontWeights(dataRange.getFontWeights()); } catch(e){}
        }
        
        // Delete original sheet from main spreadsheet
        ss.deleteSheet(sheet);
        archivedCount++;
      }
    }
    
    if (archivedCount === 0) {
      return "Aucune ancienne semaine trouvée à archiver (la semaine en cours est ignorée).";
    }
    return `${archivedCount} semaine(s) archivée(s) avec succès dans le dossier Google Drive.`;
    
  } catch (e) {
    throw new Error("Erreur lors de l'archivage: " + e.message);
  }
}

function deleteTimesheetEntry(token, rowId) {
  const sessionUser = verifySession(token);
  const username = sessionUser.username;

  try {
    const today = new Date();
    const mondayStr = getMonday(today);
    const sheetName = 'Semaine du ' + mondayStr;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.deleteRow(rowId);
    }
    return true;
  } catch (e) {
    throw new Error("Impossible de supprimer la ligne : " + e.message);
  }
}

function updateTimesheetEntry(token, updatedData) {
  const sessionUser = verifySession(token);
  const userObj = sessionUser;
  try {
    const entryDate = new Date(updatedData.entryDate);
    if (parseFloat(updatedData.duration) <= 0) {
      throw new Error("La durée doit être supérieure à 0 heure.");
    }

    const today = new Date();
    const mondayStr = getMonday(today);
    const sheetName = 'Semaine du ' + mondayStr;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const row = parseInt(updatedData.rowId);
      sheet.getRange(row, 4).setValue(updatedData.entryDate);
      sheet.getRange(row, 5).setValue(updatedData.project);
      sheet.getRange(row, 6).setValue(updatedData.task);
      sheet.getRange(row, 7).setValue(updatedData.description);
      sheet.getRange(row, 8).setValue(parseFloat(updatedData.duration));
    }
    return true;
  } catch (e) {
    throw new Error("Impossible de mettre à jour la ligne : " + e.message);
  }
}

function getEvaluationData() {
  return {
    role: 'manager',
    selfEvaluation: { achievements: 'J\'ai très bien travaillé ce mois-ci.' },
    teamEvaluations: [
      { employeeName: 'Membre Equipe', employeeEmail: 'membre@kanaga.com', selfEval: { achievements: 'Bons résultats sur les projets.' }, managerEval: null }
    ]
  };
}

function setUserLanguage(token, lang) {
  const sessionUser = verifySession(token);
 return { activeLanguage: lang }; }
function setActiveSociety(token, society) {
  const sessionUser = verifySession(token);
 return { activeSociety: society }; }
function logout() { return "https://accounts.google.com/logout"; }

// =================================================================
// 4. ODOO SYNC & AUTOMATION
// =================================================================

function setupOdooCredentials() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'ODOO_URL': 'https://kanaga-consulting.odoo.com', // Replace
    'ODOO_DB': 'kanaga-consulting',           // Replace
    'ODOO_USER': 'mhdicko@kanagaconsulting.com',   // Replace
    'ODOO_PASSWORD': 'Mahamane@2016' // Replace
  });
  Browser.msgBox("Odoo credentials have been saved successfully!");
}

function syncOdooTasks() {
  try {
    const authInfo = authenticateOdooViaJson();
    if (!authInfo) throw new Error("Could not authenticate with Odoo.");

    const projectsFR = callOdooExecuteKwJson('search_read', 'project.project', [[]], { fields: ['name', 'company_id'], context: { lang: 'fr_FR' } }, authInfo.cookie);
    const projectsEN = callOdooExecuteKwJson('search_read', 'project.project', [[]], { fields: ['name'], context: { lang: 'en_US' } }, authInfo.cookie);

    const projectENMap = new Map(projectsEN.map(p => [p.id, p.name]));
    
    const projectsForSheet = projectsFR.map(p => {
        return [ p.id, p.name, projectENMap.get(p.id) || p.name, p.company_id ? p.company_id[1] : '' ];
    });
    
    updateSheetData('Projects Odoo', projectsForSheet, 4); 

    const tasksFR = callOdooExecuteKwJson('search_read', 'project.task', [[]], { fields: ['name', 'project_id', 'company_id', 'stage_id', 'parent_id'], context: { lang: 'fr_FR' } }, authInfo.cookie);
    const tasksEN = callOdooExecuteKwJson('search_read', 'project.task', [[]], { fields: ['name', 'project_id', 'parent_id'], context: { lang: 'en_US' } }, authInfo.cookie);
    
    const taskENMap = new Map(tasksEN.map(t => [t.id, t.name]));
    
    const tasksForSheet = tasksFR.map(t => {
      let projId = t.project_id ? (Array.isArray(t.project_id) ? t.project_id[0] : t.project_id) : null;
      let parentId = t.parent_id ? (Array.isArray(t.parent_id) ? t.parent_id[0] : t.parent_id) : '';

      if (!projId && parentId) {
          const parentTask = tasksFR.find(pt => pt.id === parentId);
          if (parentTask && parentTask.project_id) {
              projId = Array.isArray(parentTask.project_id) ? parentTask.project_id[0] : parentTask.project_id;
          }
      }

      const project = projectsFR.find(p => p.id === projId);
      if (!project) return null; 
      
      const stageName = (t.stage_id && Array.isArray(t.stage_id) && t.stage_id.length > 1) ? t.stage_id[1].toLowerCase() : '';
      const isCompleted = (stageName.includes('fait') || stageName.includes('termin') || stageName.includes('done') || stageName.includes('annul') || stageName.includes('cancel') || stageName.includes('clôtur')) ? true : false;
      
      return [
        project.name, projectENMap.get(project.id) || project.name, 
        t.name, taskENMap.get(t.id) || t.name, 
        t.company_id ? (Array.isArray(t.company_id) ? t.company_id[1] : '') : '',
        isCompleted,
        t.id,
        parentId
      ];
    }).filter(Boolean);
    
    updateSheetData('ProjectTasks', tasksForSheet, 8);
    try { Browser.msgBox("Odoo Sync Complete", `${projectsForSheet.length} projects and ${tasksForSheet.length} tasks synced.`, Browser.Buttons.OK); } catch(e) {}
  } catch (e) {
    Logger.log("Trigger Sync Error: " + e.message);
    throw e;
  }
}

function triggerDebugOdoo(token) {
  const sessionUser = verifySession(token);
  try {
    const adminCheck = checkIfManager(sessionUser.username);
    if (!adminCheck.isAdmin) throw new Error("Accès refusé.");
    
    const authInfo = authenticateOdooViaJson();
    if (!authInfo) throw new Error("Could not authenticate with Odoo.");
    
    // Fetch a few tasks to see what fields Odoo 18 returns
    const tasksFR = callOdooExecuteKwJson('search_read', 'project.task', [[]], { fields: ['name', 'project_id', 'parent_id', 'child_ids'], context: { lang: 'fr_FR' } }, authInfo.cookie);
    
    let report = `Total tasks in Odoo: ${tasksFR.length}\n`;
    let subTasksCount = 0;
    
    tasksFR.forEach(t => {
       if (t.parent_id) subTasksCount++;
    });
    report += `Tasks with 'parent_id' set: ${subTasksCount}\n`;
    
    // Get up to 3 tasks that have a parent_id, else just any 3
    let sampleTasks = tasksFR.filter(t => t.parent_id).slice(0, 3);
    if (sampleTasks.length === 0) {
        sampleTasks = tasksFR.slice(0, 3);
    }
    
    const sample = sampleTasks.map(t => JSON.stringify(t)).join("\n");
    report += `\nSample Task Data:\n${sample}`;
    
    return { success: true, report: report };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateSheetData(sheetName, data, numColumns) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet '${sheetName}' not found.`);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, numColumns).setValues(data);
  }
  Logger.log(`Sheet '${sheetName}' updated with ${data.length} rows.`);
}

function runWeeklyProcess() {
  try {
    const weeklySheet = archiveStagingDataToWeeklySheet();
    if (weeklySheet) {
      pushDataFromSheetToOdoo(weeklySheet);
    }
  } catch (e) {
    Logger.log("Error in runWeeklyProcess: " + e.message);
  }
}

function archiveStagingDataToWeeklySheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const stagingSheet = ss.getSheetByName("Staging Data");
  if (!stagingSheet) return null;
  const lastRow = stagingSheet.getLastRow();
  const lastCol = stagingSheet.getLastColumn();
  if (lastRow < 2) return null;

  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const nowStr = Utilities.formatDate(now, spreadsheetTimezone, "yyyy-MM-dd HH:mm:ss");
  const today = new Date(nowStr);
  let dayOfWeek = today.getDay(); if (dayOfWeek === 0) dayOfWeek = 7;
  const daysToSubtract = dayOfWeek - 1;
  const mostRecentMonday = new Date(today.getTime());
  mostRecentMonday.setDate(today.getDate() - daysToSubtract);
  mostRecentMonday.setHours(0, 0, 0, 0);
  const sheetName = "Semaine du " + Utilities.formatDate(mostRecentMonday, spreadsheetTimezone, "yyyy-MM-dd");

  let weeklySheet = ss.getSheetByName(sheetName);
  if (weeklySheet) return weeklySheet; 
  weeklySheet = ss.insertSheet(sheetName, 0);

  const headerRange = stagingSheet.getRange(1, 1, 1, lastCol);
  const dataRange = stagingSheet.getRange(2, 1, lastRow - 1, lastCol);
  const headerValues = headerRange.getValues();
  const dataValues = dataRange.getValues();

  weeklySheet.getRange(1, 1, 1, headerValues[0].length).setValues(headerValues);
  weeklySheet.getRange(2, 1, dataValues.length, dataValues[0].length).setValues(dataValues);
  dataRange.clearContent();
  SpreadsheetApp.flush();
  return weeklySheet;
}

function archiveOldSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm Archive', 'Copy all weekly sheets to archive?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) return;

  try {
    const mainSS = SpreadsheetApp.openById(SPREADSHEET_ID);
    const mainFolder = DriveApp.getFileById(mainSS.getId()).getParents().next();
    const allSheets = mainSS.getSheets();
    const sheetsToArchive = {}; 
    const spreadsheetTimezone = mainSS.getSpreadsheetTimeZone();
    const now = new Date();
    const currentMonthYearKey = Utilities.formatDate(now, spreadsheetTimezone, "yyyy-MM");

    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const dateMatch = sheetName.match(/(\d{4}-\d{2}-\d{2})/);
      if (sheetName.startsWith('Semaine du ') && dateMatch && dateMatch[1]) {
        try {
          const dateStr = dateMatch[1];
          const sheetDate = new Date(dateStr);
          const sheetMonthYearKey = Utilities.formatDate(sheetDate, spreadsheetTimezone, "yyyy-MM");
          if (sheetMonthYearKey !== currentMonthYearKey) {
            if (!sheetsToArchive[sheetMonthYearKey]) sheetsToArchive[sheetMonthYearKey] = [];
            sheetsToArchive[sheetMonthYearKey].push(sheet);
          }
        } catch (e) {}
      }
    });

    const archiveSummary = [];
    for (const monthKey in sheetsToArchive) {
      const archiveFileName = `Timesheet Archive - ${monthKey}`;
      let archiveSS;
      let defaultSheet = null;

      const files = mainFolder.getFilesByName(archiveFileName);
      if (files.hasNext()) {
        archiveSS = SpreadsheetApp.open(files.next());
      } else {
        archiveSS = SpreadsheetApp.create(archiveFileName);
        const newFileId = archiveSS.getId();
        DriveApp.getFileById(newFileId).moveTo(mainFolder);
        defaultSheet = archiveSS.getSheetByName('Sheet1');
      }

      let sheetsArchivedThisMonth = 0;
      sheetsToArchive[monthKey].forEach(sheetToCopy => {
        const sheetName = sheetToCopy.getName();
        if (!archiveSS.getSheetByName(sheetName)) {
          sheetToCopy.copyTo(archiveSS);
          sheetsArchivedThisMonth++;
        }
      });
      
      if (defaultSheet) {
        try { archiveSS.deleteSheet(defaultSheet); } catch (e) {}
      }
      if (sheetsArchivedThisMonth > 0) archiveSummary.push(`- Copied ${sheetsArchivedThisMonth} sheets to ${archiveFileName}.`);
    }

    if (archiveSummary.length === 0) {
      ui.alert('Archive Complete', 'No old sheets were found to copy.', ui.ButtonSet.OK);
    } else {
      ui.alert('Archive Complete', archiveSummary.join('\\n'), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Archive Failed', `An error occurred: ${e.message}`, ui.ButtonSet.OK);
  }
}

function manuallyPushActiveSheetToOdoo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  if (!sheetName.startsWith("Semaine du")) {
    ui.alert("Action Canceled", "This function can only be run on a weekly timesheet sheet.", ui.ButtonSet.OK);
    return;
  }
  const response = ui.alert("Confirm Push", `Push all unsynced data from '${sheetName}' to Odoo?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    try { 
      pushDataFromSheetToOdoo(activeSheet);
      ss.toast(`Push complete.`, "Success", 10);
    } catch (e) {
      ss.toast(`Push failed: ${e.message}`, "Error", -1);
    }
  }
}

function pushDataFromSheetToOdoo(sheet) {
  try {
    if (!sheet) throw new Error("A valid sheet object was not provided.");
    
    const SYNC_STATUS_COLUMN = 10;
    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    if (allData.length < 2) return;

    const header = allData.shift();
    const rowsToSync = [];
    const originalRowNumbers = [];
    allData.forEach((row, index) => {
      if (row[0] && row[SYNC_STATUS_COLUMN - 1] !== 'Synced to Odoo') {
        rowsToSync.push(row);
        originalRowNumbers.push(index + 2); 
      }
    });

    if (rowsToSync.length === 0) return;

    const projectMap = createProjectLookupMap();
    if (projectMap.size === 0) throw new Error("Project Map is empty. Sync 'Projects Odoo' sheet.");

    const authInfo = authenticateOdooViaJson();
    const employeeMap = getOdooIdMap('hr.employee', rowsToSync, 1, authInfo.cookie, 'work_email'); 
    const companyMap = getOdooIdMap('res.company', rowsToSync, 8, authInfo.cookie); 

    const timesheetDataForOdoo = [];

    rowsToSync.forEach((row, index) => {
      const userSocietyName = row[8] ? row[8].toString().trim() : ''; 
      const projectName = row[4] ? row[4].toString().trim() : '';     
      
      const employeeId = employeeMap[row[1]]; 
      const companyId = companyMap[userSocietyName]; 

      let projectId = projectMap.get(`${projectName}|${userSocietyName}`);
      if (!projectId) {
        projectId = projectMap.get(`${projectName}|`);
      }

      if (projectId && employeeId && companyId) {
        timesheetDataForOdoo.push({
          'project_id': projectId, 
          'employee_id': employeeId, 
          'company_id': companyId, 
          'name': row[6] || '/', 
          'unit_amount': row[7], 
          'date': Utilities.formatDate(new Date(row[3]), "GMT", "yyyy-MM-dd"), 
        });
      }
    });

    if (timesheetDataForOdoo.length > 0) {
      const createdIds = callOdooExecuteKwJson('create', 'account.analytic.line', [timesheetDataForOdoo], {}, authInfo.cookie);
      const successRowNumbers = [];
      rowsToSync.forEach((row, index) => {
         const userSocietyName = row[8] ? row[8].toString().trim() : ''; 
         const projectName = row[4] ? row[4].toString().trim() : ''; 
         let projectId = projectMap.get(`${projectName}|${userSocietyName}`);
         if (!projectId) { projectId = projectMap.get(`${projectName}|`); }
         const employeeId = employeeMap[row[1]];
         const companyId = companyMap[userSocietyName];
         if (projectId && employeeId && companyId) {
           successRowNumbers.push(originalRowNumbers[index]);
         }
      });
      if (successRowNumbers.length > 0) {
         successRowNumbers.forEach(rowNum => {
           sheet.getRange(rowNum, SYNC_STATUS_COLUMN).setValue('Synced to Odoo');
         });
      }
    }
  } catch (e) {
    throw e; 
  }
}

function createProjectLookupMap() {
    const projectLookup = new Map();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const projectsSheet = ss.getSheetByName("Projects Odoo");
    
    if (!projectsSheet || projectsSheet.getLastRow() < 2) return projectLookup;
    
    const projectData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 4).getValues();
    projectData.forEach(row => {
        const odooId = row[0]; 
        const nameFr = row[1]; 
        const company = row[3] || ''; 
        if (odooId && nameFr && typeof nameFr === 'string') {
            const key = `${nameFr.trim()}|${company.toString().trim()}`;
            if (!projectLookup.has(key)) {
                projectLookup.set(key, odooId);
            }
        }
    });
    return projectLookup;
}

function getOdooIdMap(model, dataRows, columnIndex, cookie, searchField = 'name') {
  if (!dataRows || dataRows.length === 0) return {};
  const valueMap = new Map();
  dataRows.forEach(row => {
    const val = row[columnIndex];
    if (val && typeof val === 'string' && val.trim() !== '') {
      valueMap.set(val.trim(), true);
    } else if (val && (typeof val === 'number' || typeof val === 'boolean')) {
      valueMap.set(val.toString(), true);
    }
  });

  const values = Array.from(valueMap.keys());
  const idMap = {};
  if (values.length > 0) {
    try { 
      const odooData = callOdooExecuteKwJson(
        'search_read', model, [[[searchField, 'in', values]]], { fields: [searchField, 'id'] }, cookie
      );
      if (odooData) {
        odooData.forEach(item => { idMap[item[searchField]] = item.id; });
      }
    } catch(e) {}
  }
  return idMap;
}

function authenticateOdooViaJson() {
  const creds = PropertiesService.getScriptProperties().getProperties();
  const odooUrl = creds.ODOO_URL || 'https://kanaga-consulting.odoo.com';
  const odooDb = creds.ODOO_DB || 'kanaga-consulting';
  const odooUser = creds.ODOO_USER || 'mhdicko@kanagaconsulting.com';
  const odooPwd = creds.ODOO_PASSWORD || 'Mahamane@2016';
  
  const url = odooUrl + '/web/session/authenticate';
  const payload = { jsonrpc: "2.0", params: { db: odooDb, login: odooUser, password: odooPwd } };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  const httpResponse = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(httpResponse.getContentText());
  if (jsonResponse.error) throw new Error("Odoo Auth Error: " + (jsonResponse.error.data ? jsonResponse.error.data.message : jsonResponse.error.message));
  if (!jsonResponse.result || !jsonResponse.result.uid) throw new Error("Authentication failed.");
  const setCookieHeader = httpResponse.getHeaders()['Set-Cookie'] || httpResponse.getHeaders()['set-cookie'];
  const sessionCookieMatch = setCookieHeader ? setCookieHeader.match(/session_id=[^;]+/) : null;
  if (!sessionCookieMatch) throw new Error("session_id not found in cookie response.");
  return { uid: jsonResponse.result.uid, cookie: sessionCookieMatch[0] };
}

function callOdooExecuteKwJson(method, model, args, kwargs = {}, cookie) {
  const creds = PropertiesService.getScriptProperties().getProperties();
  const odooUrl = creds.ODOO_URL || 'https://kanaga-consulting.odoo.com';
  const url = odooUrl + '/web/dataset/call_kw/' + model + '/' + method;
  const payload = { jsonrpc: "2.0", params: { model, method, args, kwargs } };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true, headers: { 'Cookie': cookie } };
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();
  if (responseCode !== 200) throw new Error(`Odoo API Error: HTTP ${responseCode}`);
  const jsonResponse = JSON.parse(responseBody);
  if (jsonResponse.error) throw new Error("Odoo API Error: " + (jsonResponse.error.data ? jsonResponse.error.data.message : "Unknown error."));
  return jsonResponse.result;
}

// =================================================================
// 6. REMINDER SYSTEM 
// =================================================================

function showAdminReminderSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AdminReminderSettings').setWidth(600).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paramètres du Rappel Admin');
}
function getAdminReminderSettings() {
  const props = PropertiesService.getUserProperties();
  return {
    activation: props.getProperty('admin_reminder_activation') || 'OUI',
    destinataire: props.getProperty('admin_reminder_destinataire') || '',
    sujet: props.getProperty('admin_reminder_sujet') || '[RAPPEL] Feuilles de temps en attente',
    message: props.getProperty('admin_reminder_message') || 'Bonjour,\\n\\nLes employés suivants ont encore des feuilles de temps non soumises pour la semaine dernière :\\n\\n{liste_utilisateurs}\\n\\nMerci de faire le suivi.'
  };
}
function saveAdminReminderSettings(settings) {
  PropertiesService.getUserProperties().setProperties({
    'admin_reminder_activation': settings.activation, 'admin_reminder_destinataire': settings.destinataire,
    'admin_reminder_sujet': settings.sujet, 'admin_reminder_message': settings.message
  });
}
function sendAdminLateEntryReminder() {
  const settings = getAdminReminderSettings();
  if (settings.activation.toLowerCase() !== 'oui') return;
  const stagingSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Staging Data");
  if (!stagingSheet || stagingSheet.getLastRow() < 2) return;
  const today = new Date();
  const dayOfWeek = today.getDay();
  const startOfWeek = new Date(today.setDate(today.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1)));
  startOfWeek.setHours(0, 0, 0, 0);
  const data = stagingSheet.getDataRange().getValues();
  const lateUsers = new Set();
  for (let i = 1; i < data.length; i++) {
    const entryDate = new Date(data[i][3]);
    if (entryDate < startOfWeek) lateUsers.add(data[i][1]);
  }
  if (lateUsers.size > 0) {
    const userListString = Array.from(lateUsers).join('\\n');
    const messageFinal = settings.message.replace('{liste_utilisateurs}', userListString);
    MailApp.sendEmail(settings.destinataire, settings.sujet, messageFinal);
  }
}

function showGeneralReminderSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('GeneralReminderSettings').setWidth(600).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paramètres du Rappel Général');
}
function getGeneralReminderSettings(token) {
  const sessionUser = verifySession(token);

  const props = PropertiesService.getUserProperties();
  return {
    activation: props.getProperty('general_reminder_activation') || 'OUI',
    destinataires: props.getProperty('general_reminder_destinataires') || '',
    sujet: props.getProperty('general_reminder_sujet') || '[RAPPEL] Veuillez remplir vos feuilles de temps',
    message: props.getProperty('general_reminder_message') || 'Bonjour à tous,\\n\\nCeci est un rappel amical.'
  };
}
function saveGeneralReminderSettings(token, settings) {
  const sessionUser = verifySession(token);

  PropertiesService.getUserProperties().setProperties({
    'general_reminder_activation': settings.activation, 'general_reminder_destinataires': settings.destinataires,
    'general_reminder_sujet': settings.sujet, 'general_reminder_message': settings.message
  });
}

function getEvaluationEmailSettings(token) {
  const sessionUser = verifySession(token);

  const props = PropertiesService.getUserProperties();
  return {
    sujet: props.getProperty('eval_email_sujet') || 'Nouvelle évaluation initiée',
    message: props.getProperty('eval_email_message') || 'Bonjour,<br><br>Votre manager a initié une nouvelle évaluation pour la période <strong>{periode}</strong>.<br><br>Veuillez vous connecter au portail Kanaga pour compléter votre auto-évaluation en cliquant sur le lien suivant :<br><a href="{lien_portail}">Accéder au Portail</a><br><br>Cordialement,<br>L\'équipe Kanaga'
  };
}

function saveEvaluationEmailSettings(token, settings) {
  const sessionUser = verifySession(token);

  PropertiesService.getUserProperties().setProperties({
    'eval_email_sujet': settings.sujet,
    'eval_email_message': settings.message
  });
}
function sendGeneralTimesheetReminder() {
  const settings = getGeneralReminderSettings();
  if (settings.activation.toLowerCase() !== 'oui') return;
  if (settings.destinataires && settings.destinataires.length > 0) {
    MailApp.sendEmail({ to: settings.destinataires, subject: settings.sujet, htmlBody: settings.message, from: 'notifications@kanagaconsulting.com' });
  }
}

function syncOdooEmployees() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName("Utilisateurs");
    if (!userSheet) throw new Error("Sheet 'Utilisateurs' not found.");

    const oldData = userSheet.getDataRange().getValues();
    const passwordMap = new Map();
    const companyMap = new Map();
    const birthdateMap = new Map(); 
    const roleMap = new Map();

    for (let i = 1; i < oldData.length; i++) {
      const email = oldData[i][0]; 
      if (email) {
        const emailKey = email.toString().trim();
        const password = oldData[i][1]; 
        const company = oldData[i][3];  
        const birthdate = oldData[i][4]; 
        const role = oldData[i][5];
        // Fixing variable name below
        if (password) passwordMap.set(emailKey, password.toString());
        if (company) companyMap.set(emailKey, company.toString());
        if (birthdate) { 
          let birthdateString = (birthdate instanceof Date) ? Utilities.formatDate(birthdate, Session.getScriptTimeZone(), "yyyy-MM-dd") : birthdate.toString();
          birthdateMap.set(emailKey, birthdateString);
        }
        if (oldData[i][5]) roleMap.set(emailKey, oldData[i][5].toString());
      }
    }

    const authInfo = authenticateOdooViaJson();
    if (!authInfo) throw new Error("Could not authenticate with Odoo.");

    const odooCompanies = callOdooExecuteKwJson('search_read', 'res.company', [[]], { fields: ['name'] }, authInfo.cookie);
    const odooCompanyMap = new Map(odooCompanies.map(c => [c.id, c.name]));

    const employeeFields = ['name', 'company_id', 'work_email', 'user_id', 'birthday'];
    const employeesOdoo = callOdooExecuteKwJson('search_read', 'hr.employee', [ [['active', '=', true]] ], { fields: employeeFields }, authInfo.cookie);

    const userIds = employeesOdoo.map(emp => emp.user_id ? emp.user_id[0] : null).filter(Boolean);
    let userLoginMap = new Map();
    if (userIds.length > 0) {
      const usersOdoo = callOdooExecuteKwJson('search_read', 'res.users', [ [['id', 'in', userIds]] ], { fields: ['login'] }, authInfo.cookie);
      usersOdoo.forEach(user => { userLoginMap.set(user.id, user.login); });
    }

    const newData = [];
    const rowsToColorRed = []; 
    
    for (const employee of employeesOdoo) {
      const fullName = employee.name;
      let primaryEmail = employee.user_id ? userLoginMap.get(employee.user_id[0]) : null;
      if (!primaryEmail && employee.work_email) primaryEmail = employee.work_email;
      if (!primaryEmail) continue;

      const username = primaryEmail.toString().trim();
      let mainCompanyName = employee.company_id ? (odooCompanyMap.get(employee.company_id[0]) || employee.company_id[1]) : '';
      const password = passwordMap.get(username) || 'password';
      const company = companyMap.get(username) || mainCompanyName;
      
      const newBirthday = employee.birthday || ''; 
      const oldBirthday = birthdateMap.get(username) || ''; 
      if (newBirthday !== oldBirthday) rowsToColorRed.push(newData.length);
      
      const role = roleMap.get(username) || 'Utilisateur';

      newData.push([ username, password, fullName, company, newBirthday, role ]);
    }

    const lastRow = userSheet.getLastRow();
    if (lastRow > 1) userSheet.getRange(2, 1, lastRow - 1, 6).clearContent().setBackground(null); 
    if (newData.length > 0) userSheet.getRange(2, 1, newData.length, 6).setValues(newData);
    
    if (rowsToColorRed.length > 0) {
      const rangesToColor = rowsToColorRed.map(rowIndex => `E${rowIndex + 2}`);
      userSheet.getRangeList(rangesToColor).setBackground("#FF0000"); 
    }
    try { Browser.msgBox("Odoo Employee Sync Complete", `${newData.length} employees synced.`, Browser.Buttons.OK); } catch (e) {}
  } catch (e) {
    try { Browser.msgBox("Employee Sync Error", "Details: ".concat(e.message), Browser.Buttons.OK); } catch (err) {}
  }
}

// =================================================================
// 5. NATIVE EVALUATION SYSTEM (Modèle en Sablier / Y)
// =================================================================

function getEvaluationConfig(token) {
  const sessionUser = verifySession(token);

  const defaultScale = "PI, PA, CA, PS, PE, N/A";
  const defaultConfig = {
    fondamentales: [
      { text: "1. Professionnalisme et éthique", description: "Normes, confidentialité, intégrité, ponctualité, présentation et exemplarité.", type: "scale", options: defaultScale, page: 1 },
      { text: "2. Communication Orale et Écrite", description: "Clarté, concision, écoute active, adaptation du message aux interlocuteurs.", type: "scale", options: defaultScale, page: 1 },
      { text: "3. Travail d'équipe et Collaboration", description: "Contribution positive, esprit d'entraide, partage d'informations et soutien aux collègues.", type: "scale", options: defaultScale, page: 1 },
      { text: "4. Organisation et Gestion du Temps", description: "Planification, priorisation des tâches, respect strict des échéances.", type: "scale", options: defaultScale, page: 1 },
      { text: "5. Initiative et Proactivité", description: "Capacité à anticiper les besoins, formuler des propositions et résoudre les blocages.", type: "scale", options: defaultScale, page: 1 },
      { text: "6. Adaptabilité et Apprentissage Continu", description: "Flexibilité face au changement, curiosité intellectuelle, volonté de perfectionnement.", type: "scale", options: defaultScale, page: 1 },
      { text: "7. Compréhension du Contexte Local", description: "Maîtrise de l'environnement économique, réglementaire et culturel des interventions.", type: "scale", options: defaultScale, page: 1 },
      { text: "8. Capacité d'Innovation", description: "Recherche de méthodes innovantes, adoption de nouveaux outils et création de valeur.", type: "scale", options: defaultScale, page: 1 }
    ],
    "Junior - Secteur privé": [
      { text: "1. Compréhension des enjeux business privés", description: "Capacité à cerner le modèle économique, la rentabilité et les priorités du client privé.", type: "scale", options: defaultScale, page: 1 },
      { text: "2. Rigueur d'exécution et qualité des livrables", description: "Précision, respect des formats attendus, souci du détail et respect des délais clients.", type: "scale", options: defaultScale, page: 1 },
      { text: "3. Réactivité et orientation service client", description: "Disponibilité, écoute des attentes opérationnelles et relation client de confiance.", type: "scale", options: defaultScale, page: 1 },
      { text: "4. Maîtrise des outils bureautiques et digitaux", description: "Utilisation efficace d'Excel, PowerPoint, ERP et logiciels métiers.", type: "scale", options: defaultScale, page: 2 },
      { text: "5. Polyvalence et esprit d'équipe en mission", description: "Capacité à s'intégrer rapidement sur des projets variés du secteur privé.", type: "scale", options: defaultScale, page: 2 }
    ],
    "Assistant - Secteur public": [
      { text: "1. Connaissance des procédures administratives publiques", description: "Respect des circuits de validation, règles de gestion et procédures du secteur public.", type: "scale", options: defaultScale, page: 1 },
      { text: "2. Conformité réglementaire et archivage", description: "Rigueur documentaire, traçabilité des actes et conformité aux règles des bailleurs.", type: "scale", options: defaultScale, page: 1 },
      { text: "3. Rédaction administrative et comptes-rendus", description: "Qualité rédactionnelle formelle adaptée à l'administration publique.", type: "scale", options: defaultScale, page: 1 },
      { text: "4. Suivi d'exécution budgétaire / marchés", description: "Vérification des pièces justificatives, suivi des engagements et mandats.", type: "scale", options: defaultScale, page: 2 },
      { text: "5. Sens du service public et relation usagers/bailleurs", description: "Déontologie, neutralité et collaboration efficace avec les partenaires institutionnels.", type: "scale", options: defaultScale, page: 2 }
    ],
    "Consultant": [
      { text: "1. Analyse et Résolution de Problèmes", description: "Identifier les enjeux stratégiques, analyser les données et concevoir des solutions pragmatiques.", type: "scale", options: defaultScale, page: 1 },
      { text: "2. Gestion de Projet / Mission", description: "Cadrage, planification, pilotage des livrables, gestion des risques et respect du budget.", type: "scale", options: defaultScale, page: 1 },
      { text: "3. Qualité et Pertinence des Livrables", description: "Excellence rédactionnelle, valeur ajoutée des recommandations et professionnalisme.", type: "scale", options: defaultScale, page: 1 },
      { text: "4. Relation Client et Posture Conseil", description: "Compréhension approfondie des besoins, gestion de la satisfaction et écoute active.", type: "scale", options: defaultScale, page: 2 },
      { text: "5. Contribution au Développement Commercial", description: "Détection d'opportunités, participation aux propositions commerciales et foisonnement.", type: "scale", options: defaultScale, page: 2 }
    ],
    "Comptable": [
      { text: "1. Rigueur et Exactitude Comptable", description: "Précision dans la saisie, révision des comptes et préparation des états financiers.", type: "scale", options: defaultScale, page: 1 },
      { text: "2. Respect des Délais et Procédures Légales", description: "Clôtures périodiques, déclarations fiscales et sociales dans les délais prescrits.", type: "scale", options: defaultScale, page: 1 },
      { text: "3. Analyse et Contrôle de Gestion", description: "Suivi budgétaire, analyse des écarts et propositions d'optimisation financière.", type: "scale", options: defaultScale, page: 1 },
      { text: "4. Maîtrise des Logiciels Financiers", description: "Utilisation experte des logiciels comptables, ERP et modélisation sous tableur.", type: "scale", options: defaultScale, page: 2 },
      { text: "5. Veille Réglementaire et Fiscale", description: "Maintien à jour des connaissances en droit fiscal, comptable et social.", type: "scale", options: defaultScale, page: 2 }
    ],
    "Chef de Projet": [
      { text: "1. Cadrage stratégique et définition des objectifs", description: "Compréhension des attentes clients, note de cadrage et alignement des parties prenantes.", type: "scale", options: defaultScale, page: 1 },
      { text: "2. Planification et gestion des ressources", description: "Établissement du rétroplanning, allocation des compétences et anticipation des charges.", type: "scale", options: defaultScale, page: 1 },
      { text: "3. Maîtrise des délais, coûts et livrables", description: "Respect des jalons, contrôle budgétaire et garantie de la qualité des livrables.", type: "scale", options: defaultScale, page: 1 },
      { text: "4. Leadership et animation d'équipe", description: "Coordination des consultants, motivation, délégation efficace et feedback constructif.", type: "scale", options: defaultScale, page: 2 },
      { text: "5. Gestion des risques et résolution de crises", description: "Identification proactive des risques, plans d'atténuation et arbitrage.", type: "scale", options: defaultScale, page: 2 },
      { text: "6. Posture conseil et relation de confiance client", description: "Communication executive, comités de pilotage et fidélisation client.", type: "scale", options: defaultScale, page: 2 }
    ],
    "Auditeur": [
      { text: "1. Maîtrise des normes et référentiels d'audit", description: "Application des standards comptables/financiers et réglementations applicables.", type: "scale", options: defaultScale, page: 1 },
      { text: "2. Évaluation du contrôle interne et cartographie des risques", description: "Analyse des processus, détection des faiblesses et zones de vulnérabilité.", type: "scale", options: defaultScale, page: 1 },
      { text: "3. Rigueur d'échantillonnage et vérification des données", description: "Collecte des preuves d'audit, exhaustivité des tests de conformité.", type: "scale", options: defaultScale, page: 1 },
      { text: "4. Qualité rédactionnelle des constats et rapports", description: "Précision des faits, clarté de la synthèse et objectivité.", type: "scale", options: defaultScale, page: 2 },
      { text: "5. Recommandations pragmatiques à forte valeur ajoutée", description: "Pertinence opérationnelle, faisabilité et hiérarchisation des actions.", type: "scale", options: defaultScale, page: 2 },
      { text: "6. Posture professionnelle et indépendance", description: "Esprit critique constructif, neutralité et respect de la déontologie.", type: "scale", options: defaultScale, page: 2 }
    ],
    "conclusion": [
      { text: "Réalisation des Objectifs de la période écoulée", description: "Détaillez les objectifs fixés, résultats atteints, faits marquants et difficultés rencontrées.", type: "text", options: "", page: 1 },
      { text: "1. Principaux Points Forts de l'Employé(e)", description: "Compétences clés démontrées, succès et contributions notables.", type: "text", options: "", page: 1 },
      { text: "2. Axes d'Amélioration Prioritaires", description: "Domaines nécessitant un perfectionnement ou un accompagnement particulier.", type: "text", options: "", page: 1 },
      { text: "3. Appréciation de la Performance Globale", description: "Appréciation synthétique de la performance générale sur la période.", type: "scale", options: "PI, PA, CA, PS, PE", page: 1 },
      { text: "VI. 1. Besoins en Formation", description: "Formations techniques, managériales ou linguistiques souhaitées pour la progression.", type: "text", options: "", page: 2 },
      { text: "VI. 2. Objectifs SMART pour la Prochaine Période", description: "Objectifs Spécifiques, Mesurables, Atteignables, Réalistes et Temporellement définis.", type: "text", options: "", page: 2 },
      { text: "VI. 3. Aspirations Professionnelles", description: "Évolution de carrière envisagée, souhaits de mobilité ou nouvelles responsabilités.", type: "text", options: "", page: 2 },
      { text: "VII. Commentaires de l'Employé(e) Évalué(e)", description: "Remarques complémentaires du collaborateur.", type: "text", options: "", page: 2 }
    ]
  };

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('EvaluationQuestions');

    // Si la feuille n'existe pas ou est vide, l'initialiser avec les questions par défaut
    if (!sheet || sheet.getLastRow() <= 1) {
      if (!sheet) {
        sheet = ss.insertSheet('EvaluationQuestions');
      } else {
        sheet.clear();
      }
      const headers = ['Profil', 'Base Associée', 'Page', 'Ordre', 'Question', 'Description', 'Type', 'Options'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');

      const initialRows = [];
      for (let prof in defaultConfig) {
        const baseAssociee = (prof.toLowerCase() === 'fondamentales' || prof.toLowerCase() === 'conclusion') ? '' : 'fondamentales';
        defaultConfig[prof].forEach((q, idx) => {
          initialRows.push([
            prof,
            baseAssociee,
            q.page || 1,
            idx + 1,
            q.text || '',
            q.description || '',
            q.type || 'scale',
            q.options || defaultScale
          ]);
        });
      }
      if (initialRows.length > 0) {
        sheet.getRange(2, 1, initialRows.length, headers.length).setValues(initialRows);
      }
      return defaultConfig;
    }

    // Lecture dynamique depuis la feuille EvaluationQuestions
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0].map(h => String(h || '').trim().toLowerCase());

    const colProfil = headerRow.indexOf('profil');
    let colPage = headerRow.indexOf('page');
    const colOrdre = headerRow.indexOf('ordre');
    const colQuestion = headerRow.indexOf('question');
    const colDesc = headerRow.indexOf('description');
    const colType = headerRow.indexOf('type');
    const colOptions = headerRow.indexOf('options');

    // Si la colonne 'Page' n'existe pas encore dans la feuille Google Sheets, l'ajouter immédiatement
    if (colPage === -1) {
      try {
        sheet.insertColumnAfter(2); // Insère la colonne en 3ème position (entre 'Base Associée' et 'Ordre')
        sheet.getRange(1, 3).setValue('Page').setFontWeight('bold').setBackground('#f3f3f3');
        if (data.length > 1) {
          const pageValues = [];
          for (let r = 1; r < data.length; r++) {
            const prof = String(data[r][colProfil !== -1 ? colProfil : 0] || '').trim().toLowerCase();
            const ord = parseInt(data[r][colOrdre !== -1 ? colOrdre : 2], 10) || 1;
            let pNum = 1;
            if (prof.includes('junior') || prof.includes('assistant') || prof.includes('consultant') || prof.includes('comptable') || prof.includes('chef') || prof.includes('audit')) {
              pNum = (ord > 3) ? 2 : 1;
            } else if (prof === 'conclusion') {
              pNum = (ord > 4) ? 2 : 1;
            }
            pageValues.push([pNum]);
          }
          sheet.getRange(2, 3, pageValues.length, 1).setValues(pageValues);
        }
        
        // Recharger les données et en-têtes suite à l'insertion de la colonne
        data = sheet.getDataRange().getValues();
        headerRow = data[0].map(h => String(h || '').trim().toLowerCase());
        colPage = headerRow.indexOf('page');
      } catch(migErr) {
        Logger.log("Erreur migration colonne Page: " + migErr.message);
      }
    }

    const config = {};

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const prof = String(row[colProfil !== -1 ? colProfil : 0] || '').trim();
      if (!prof) continue;

      if (!config[prof]) {
        config[prof] = [];
      }

      const page = colPage !== -1 ? (parseInt(row[colPage], 10) || 1) : 1;
      const order = colOrdre !== -1 ? (parseInt(row[colOrdre], 10) || (config[prof].length + 1)) : (config[prof].length + 1);
      const text = String(colQuestion !== -1 ? row[colQuestion] : (row[3] || '')).trim();
      const desc = String(colDesc !== -1 ? row[colDesc] : (row[4] || '')).trim();
      const type = String(colType !== -1 ? row[colType] : (row[5] || 'scale')).trim() || 'scale';
      const options = String(colOptions !== -1 ? row[colOptions] : (row[6] || defaultScale)).trim() || defaultScale;

      config[prof].push({
        text: text,
        description: desc,
        type: type,
        options: options,
        page: page,
        order: order
      });
    }

    // Auto-fusion de la conclusion si absente de la feuille
    if (!config['conclusion'] || config['conclusion'].length === 0) {
      config['conclusion'] = defaultConfig['conclusion'];
    } else if (config['conclusion'] && config['conclusion'].length >= 5) {
      // Si toutes les questions de la conclusion sont sur la page 1, rétablir les 2 pages standard
      const allPage1 = config['conclusion'].every(q => (parseInt(q.page, 10) || 1) === 1);
      if (allPage1) {
        config['conclusion'].forEach((q, idx) => {
          q.page = (idx >= 4) ? 2 : 1;
        });
      }
    }

    // Tri de chaque profil par la colonne Ordre
    for (let p in config) {
      config[p].sort((a, b) => (a.order || 0) - (b.order || 0));
    }

    config._stepTitles = getEvaluationStepTitles();
    return config;
  } catch(e) {
    Logger.log("Erreur lecture EvaluationQuestions: " + e.message);
    defaultConfig._stepTitles = getEvaluationStepTitles();
    return defaultConfig;
  }
}

function saveEvaluationConfig(token, config) {
  const sessionUser = verifySession(token);
  const userRole = ((sessionUser && sessionUser.role) || '').toLowerCase().trim();
  if (userRole !== 'admin' && userRole !== 'manager' && userRole !== 'directeur') {
    throw new Error("Action non autorisée. Seuls les administrateurs et managers peuvent modifier les profils et questions.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('EvaluationQuestions');
  if (!sheet) {
    sheet = ss.insertSheet('EvaluationQuestions');
  }

  sheet.clear();
  const headers = ['Profil', 'Base Associée', 'Page', 'Ordre', 'Question', 'Description', 'Type', 'Options'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');

  const rowsToAdd = [];
  const defaultScale = "PI, PA, CA, PS, PE, N/A";

  for (let profile in config) {
    if (profile.startsWith('_') || !Array.isArray(config[profile])) continue;
    const qList = config[profile];
    const baseAssociee = (profile.toLowerCase() === 'fondamentales' || profile.toLowerCase() === 'conclusion') ? '' : 'fondamentales';
    if (Array.isArray(qList)) {
      qList.forEach((q, idx) => {
        rowsToAdd.push([
          profile,
          baseAssociee,
          parseInt(q.page, 10) || 1,
          idx + 1,
          q.text || '',
          q.description || '',
          q.type || 'scale',
          q.options || defaultScale
        ]);
      });
    }
  }

  if (rowsToAdd.length > 0) {
    sheet.getRange(2, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
  }

  return { success: true, count: rowsToAdd.length };
}

function getEvaluationStepTitles() {
  const defaultTitles = {
    step1: "1. Entonnoir (Infos)",
    step2: "2. Tronc Commun",
    step3: "3. Aiguillage Profil",
    step4: "4. Réunion Finale"
  };
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('EVALUATION_STEP_TITLES');
    if (raw) {
      const parsed = JSON.parse(raw);
      return {
        step1: parsed.step1 || defaultTitles.step1,
        step2: parsed.step2 || defaultTitles.step2,
        step3: parsed.step3 || defaultTitles.step3,
        step4: parsed.step4 || defaultTitles.step4
      };
    }
  } catch(e) {
    Logger.log("Erreur lecture step titles: " + e.message);
  }
  return defaultTitles;
}

function saveEvaluationStepTitles(token, titles) {
  const sessionUser = verifySession(token);
  const userRole = ((sessionUser && sessionUser.role) || '').toLowerCase().trim();
  if (userRole !== 'admin' && userRole !== 'manager' && userRole !== 'directeur') {
    throw new Error("Action non autorisée. Seuls les administrateurs et managers peuvent modifier les intitulés des étapes.");
  }

  const cleanTitles = {
    step1: String((titles && titles.step1) || "1. Entonnoir (Infos)").trim(),
    step2: String((titles && titles.step2) || "2. Tronc Commun").trim(),
    step3: String((titles && titles.step3) || "3. Aiguillage Profil").trim(),
    step4: String((titles && titles.step4) || "4. Réunion Finale").trim()
  };

  PropertiesService.getScriptProperties().setProperty('EVALUATION_STEP_TITLES', JSON.stringify(cleanTitles));
  return { success: true, titles: cleanTitles };
}

function initiateEvaluationsBatch(token, formData) {
  const sessionUser = verifySession(token);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Evaluations');
    if (!sheet) {
      sheet = ss.insertSheet('Evaluations');
      sheet.appendRow(['Timestamp', 'Statut', 'Période', 'Employé', 'Département', 'Date Entrée', 'Évaluateur', 'Poste Évaluateur', 'Date Entretien', 'Profil', 'Fondamentales', 'Spécifiques', 'Objectifs Passés', 'Points Forts', 'Améliorations', 'Note Globale', 'Formation', 'Objectifs SMART', 'Aspirations', 'Commentaires Employé', 'Mgr Name', 'Mgr Fondamentales', 'Mgr Spécifiques', 'Mgr Points Forts', 'Mgr Améliorations', 'Mgr Note Globale', 'Mgr Formation', 'Mgr Objectifs SMART', 'Mgr Commentaires', 'Employee Email', 'PDF URL', 'Principal Evaluator Email', 'Principal Evaluator Name', 'Secondary Evaluators', 'Secondary Fondamentales', 'Secondary Specifiques', 'Secondary Synthese']);
      sheet.getRange("A1:AK1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const maxRequiredCol = 37;
    if (sheet.getMaxColumns() < maxRequiredCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxRequiredCol - sheet.getMaxColumns());
    }

    const timestamp = new Date().toLocaleString('fr-FR');
    const employees = formData.employees || [];
    
    const principalEmail = (formData.principalEvaluator && formData.principalEvaluator.email) || (sessionUser ? sessionUser.username : '');
    const principalName = (formData.principalEvaluator && formData.principalEvaluator.name) || (sessionUser ? sessionUser.fullName : formData.evaluatorName) || 'Évaluateur Principal';
    
    const secondaryList = (formData.secondaryEvaluators || []).map(s => ({
      email: s.email,
      name: s.name,
      status: 'En attente',
      submittedAt: ''
    }));

    employees.forEach(emp => {
        const rowData = [
          timestamp,
          'Initiée',
          formData.period || '',
          emp.name || '',
          '', // dept
          '', // entryDate
          principalName,
          '', // evaluatorJob
          '', // interviewDate
          formData.profile || '',
          '', '', '', '', '', '', '', '', '', '', 
          '', '', '', '', '', '', '', '', '', 
          emp.email || '',
          '', // PDF URL (col 31)
          principalEmail, // col 32
          principalName,  // col 33
          JSON.stringify(secondaryList), // col 34
          JSON.stringify([]), // col 35 (Secondary Fondamentales)
          JSON.stringify([]), // col 36 (Secondary Specifiques)
          JSON.stringify([])  // col 37 (Secondary Synthese)
        ];

        sheet.appendRow(rowData);
        
        // Envoi d'email de notification au collaborateur
        if (emp.email) {
           try {
             const emailSettings = getEvaluationEmailSettings();
             const portailUrl = ScriptApp.getService().getUrl();
             const subject = emailSettings.sujet.replace('{periode}', formData.period || "");
             const bodyHtml = emailSettings.message
                                .replace(/{periode}/g, formData.period || "")
                                .replace(/{lien_portail}/g, portailUrl);
                                
             MailApp.sendEmail({
               to: emp.email,
               subject: subject,
               htmlBody: bodyHtml
             });
           } catch(mailErr) {
             console.error("Erreur envoi email: " + mailErr.message);
           }
        }
    });
    
    return { success: true };
  } catch (e) {
    throw new Error("Erreur lors de l'initiation de l'évaluation: " + e.message);
  }
}

function submitSelfEvaluation(token, rowId, formData) {
  let userEmail = '';
  if (token && typeof token === 'string' && token.length > 10) {
    try {
      const sessionUser = verifySession(token);
      userEmail = sessionUser.username;
    } catch(e) {}
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error("Feuille Evaluations introuvable.");
    
    const row = parseInt(rowId);
    const maxRequiredCol = 37;
    if (sheet.getMaxColumns() < maxRequiredCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxRequiredCol - sheet.getMaxColumns());
    }

    // Déterminer le prochain statut selon la présence d'évaluateurs secondaires
    let nextStatus = 'Attente Évaluateur Principal';
    const secondaryStr = sheet.getRange(row, 34).getValue();
    let secondaryList = [];
    try {
      if (secondaryStr) secondaryList = JSON.parse(secondaryStr);
    } catch(e) {}

    if (Array.isArray(secondaryList) && secondaryList.length > 0) {
      nextStatus = 'Attente Évaluateur Secondaire';
    }

    sheet.getRange(row, 2).setValue(nextStatus);
    sheet.getRange(row, 11).setValue(JSON.stringify(formData.fondamentales || []));
    sheet.getRange(row, 12).setValue(JSON.stringify(formData.specifiques || []));
    sheet.getRange(row, 13).setValue(formData.pastGoals || '');
    sheet.getRange(row, 14).setValue(formData.strengths || '');
    sheet.getRange(row, 15).setValue(formData.improvements || '');
    sheet.getRange(row, 16).setValue(formData.globalRating || '');
    sheet.getRange(row, 17).setValue(formData.training || '');
    sheet.getRange(row, 18).setValue(formData.smartGoals || '');
    sheet.getRange(row, 19).setValue(formData.aspirations || '');
    sheet.getRange(row, 20).setValue(formData.empComments || '');

    return { success: true, nextStatus: nextStatus };
  } catch (e) {
    throw new Error("Erreur lors de l'enregistrement de l'auto-évaluation: " + e.message);
  }
}

function submitSecondaryEvaluation(token, rowId, formData) {
  let evaluatorName = '';
  let evaluatorEmail = '';
  if (token && typeof token === 'string' && token.length > 10) {
    try {
      const sessionUser = verifySession(token);
      evaluatorName = sessionUser.fullName || sessionUser.username;
      evaluatorEmail = sessionUser.username;
    } catch(e) {}
  }
  if (!evaluatorName) evaluatorName = formData.evaluatorName || 'Évaluateur Secondaire';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error("Feuille Evaluations introuvable.");
    
    const row = parseInt(rowId);
    const maxRequiredCol = 37;
    if (sheet.getMaxColumns() < maxRequiredCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxRequiredCol - sheet.getMaxColumns());
    }

    // 1. Mettre à jour le statut de l'évaluateur secondaire dans la liste
    let secondaryList = [];
    try {
      const secStr = sheet.getRange(row, 34).getValue();
      if (secStr) secondaryList = JSON.parse(secStr);
    } catch(e) {}

    let matched = false;
    secondaryList.forEach(s => {
      if ((evaluatorEmail && s.email && s.email.toLowerCase() === evaluatorEmail.toLowerCase()) ||
          (evaluatorName && s.name && s.name.toLowerCase() === evaluatorName.toLowerCase())) {
        s.status = 'Complété';
        s.submittedAt = new Date().toLocaleString('fr-FR');
        matched = true;
      }
    });
    if (!matched && evaluatorEmail) {
      secondaryList.push({
        email: evaluatorEmail,
        name: evaluatorName,
        status: 'Complété',
        submittedAt: new Date().toLocaleString('fr-FR')
      });
    }

    // 2. Enregistrer dans secondaryFondamentales
    let secFonda = [];
    try {
      const sfStr = sheet.getRange(row, 35).getValue();
      if (sfStr) secFonda = JSON.parse(sfStr);
    } catch(e) {}
    secFonda = secFonda.filter(entry => entry.evaluatorEmail !== evaluatorEmail);
    secFonda.push({
      evaluatorEmail: evaluatorEmail,
      evaluatorName: evaluatorName,
      submittedAt: new Date().toLocaleString('fr-FR'),
      responses: formData.fondamentales || []
    });

    // 3. Enregistrer dans secondarySpecifiques
    let secSpec = [];
    try {
      const ssStr = sheet.getRange(row, 36).getValue();
      if (ssStr) secSpec = JSON.parse(ssStr);
    } catch(e) {}
    secSpec = secSpec.filter(entry => entry.evaluatorEmail !== evaluatorEmail);
    secSpec.push({
      evaluatorEmail: evaluatorEmail,
      evaluatorName: evaluatorName,
      submittedAt: new Date().toLocaleString('fr-FR'),
      responses: formData.specifiques || []
    });

    // 4. Enregistrer dans secondarySynthese
    let secSynth = [];
    try {
      const syStr = sheet.getRange(row, 37).getValue();
      if (syStr) secSynth = JSON.parse(syStr);
    } catch(e) {}
    secSynth = secSynth.filter(entry => entry.evaluatorEmail !== evaluatorEmail);
    secSynth.push({
      evaluatorEmail: evaluatorEmail,
      evaluatorName: evaluatorName,
      submittedAt: new Date().toLocaleString('fr-FR'),
      strengths: formData.strengths || '',
      improvements: formData.improvements || '',
      globalRating: formData.globalRating || '',
      training: formData.training || '',
      smartGoals: formData.smartGoals || '',
      comments: formData.empComments || formData.comments || ''
    });

    // 5. Vérifier si tous les évaluateurs secondaires ont terminé
    const allCompleted = secondaryList.length === 0 || secondaryList.every(s => s.status === 'Complété');
    const nextStatus = allCompleted ? 'Attente Évaluateur Principal' : 'Attente Évaluateur Secondaire';

    sheet.getRange(row, 2).setValue(nextStatus);
    sheet.getRange(row, 34).setValue(JSON.stringify(secondaryList));
    sheet.getRange(row, 35).setValue(JSON.stringify(secFonda));
    sheet.getRange(row, 36).setValue(JSON.stringify(secSpec));
    sheet.getRange(row, 37).setValue(JSON.stringify(secSynth));

    return { success: true, nextStatus: nextStatus, allCompleted: allCompleted };
  } catch (e) {
    throw new Error("Erreur lors de l'enregistrement de l'évaluation secondaire: " + e.message);
  }
}

function advanceToPrincipalEvaluation(token, rowId) {
  const sessionUser = verifySession(token);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error("Feuille Evaluations introuvable.");
    const row = parseInt(rowId);
    sheet.getRange(row, 2).setValue('Attente Évaluateur Principal');
    return { success: true };
  } catch(e) {
    throw new Error("Erreur advanceToPrincipalEvaluation: " + e.message);
  }
}

function saveSelfEvaluationDraft(token, rowId, formData) {
  const sessionUser = verifySession(token);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error("Feuille Evaluations introuvable.");
    
    const row = parseInt(rowId);
    const currentStatus = sheet.getRange(row, 2).getValue();
    if (currentStatus === 'Initiée') {
      sheet.getRange(row, 2).setValue('Brouillon Employé');
    }

    sheet.getRange(row, 11).setValue(JSON.stringify(formData.fondamentales || []));
    sheet.getRange(row, 12).setValue(JSON.stringify(formData.specifiques || []));
    sheet.getRange(row, 13).setValue(formData.pastGoals || '');
    sheet.getRange(row, 14).setValue(formData.strengths || '');
    sheet.getRange(row, 15).setValue(formData.improvements || '');
    sheet.getRange(row, 16).setValue(formData.globalRating || '');
    sheet.getRange(row, 17).setValue(formData.training || '');
    sheet.getRange(row, 18).setValue(formData.smartGoals || '');
    sheet.getRange(row, 19).setValue(formData.aspirations || '');
    sheet.getRange(row, 20).setValue(formData.empComments || '');

    return { success: true };
  } catch (e) {
    throw new Error("Erreur lors de la sauvegarde du brouillon de l'auto-évaluation: " + e.message);
  }
}

function generateEvaluationPDF(rowId) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    const data = sheet.getRange(rowId, 1, 1, 37).getValues()[0];
    
    const employeeName = data[3] || '';
    const period = data[2] || '';
    
    const docName = `Evaluation_${employeeName}_${period}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    
    // Page margins and font
    body.setMarginLeft(40);
    body.setMarginRight(40);
    body.setMarginTop(40);
    body.setMarginBottom(40);
    
    // Title
    const title = body.appendParagraph("Fiche d'Évaluation");
    title.setHeading(DocumentApp.ParagraphHeading.TITLE);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.setForegroundColor('#8b5a2b'); // Brown/Gold color
    
    body.appendParagraph(""); // Spacer
    
    // Noms des évaluateurs secondaires
    let secNames = '';
    try {
      const secList = JSON.parse(data[33] || '[]');
      secNames = secList.map(s => s.name || s.email).join(', ');
    } catch(e) {}

    // Top Info Header Table
    const infoTable = body.appendTable([
      ["Statut:", data[1] || '', "Période:", period],
      ["Employé:", employeeName, "Département:", data[4] || ''],
      ["Évaluateur Principal:", data[32] || data[20] || data[6] || '', "Profil:", data[9] || ''],
      ["Évaluateur(s) Secondaire(s):", secNames || 'Aucun', "", ""]
    ]);
    
    // Format Header Table
    for (let i = 0; i < infoTable.getNumRows(); i++) {
      const row = infoTable.getRow(i);
      row.getCell(0).getChild(0).asParagraph().setAttributes({ [DocumentApp.Attribute.BOLD]: true });
      row.getCell(2).getChild(0).asParagraph().setAttributes({ [DocumentApp.Attribute.BOLD]: true });
      for (let j = 0; j < row.getNumCells(); j++) {
        row.getCell(j).setPaddingTop(5).setPaddingBottom(5);
      }
    }
    
    body.appendParagraph(""); // Spacer
    
    // Parse secondary synth
    let secSynthList = [];
    try { secSynthList = JSON.parse(data[36] || '[]'); } catch(e) {}
    let secStrengths = secSynthList.map(s => (s.evaluatorName ? s.evaluatorName + " : " : "") + (s.strengths || '-')).join("\n");
    let secImprovements = secSynthList.map(s => (s.evaluatorName ? s.evaluatorName + " : " : "") + (s.improvements || '-')).join("\n");
    let secGlobal = secSynthList.map(s => (s.evaluatorName ? s.evaluatorName + " : " : "") + (s.globalRating || '-')).join(", ");
    let secComments = secSynthList.map(s => (s.evaluatorName ? s.evaluatorName + " : " : "") + (s.comments || '-')).join("\n");

    // Synthèse Globale Section
    const synthHeading = body.appendParagraph("Synthèse Globale");
    synthHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    synthHeading.setAttributes({ [DocumentApp.Attribute.BOLD]: true });
    
    const synthTable = body.appendTable([
      ["Rubrique", "Auto-évaluation (Employé)", "Évaluation(s) Secondaire(s)", "Évaluation Principale"],
      ["Note Globale", data[15] || '-', secGlobal || '-', data[25] || '-'],
      ["Points Forts", data[13] || '-', secStrengths || '-', data[23] || '-'],
      ["Axes d'Amélioration", data[14] || '-', secImprovements || '-', data[24] || '-'],
      ["Besoins en Formation", data[16] || '-', '-', data[26] || '-'],
      ["Objectifs SMART", data[17] || '-', '-', data[27] || '-'],
      ["Commentaires", data[19] || '-', secComments || '-', data[28] || '-']
    ]);
    
    // Format Synth Table
    const synthHeader = synthTable.getRow(0);
    for (let j = 0; j < synthHeader.getNumCells(); j++) {
      synthHeader.getCell(j).getChild(0).asParagraph().setAttributes({ [DocumentApp.Attribute.BOLD]: true });
      synthHeader.getCell(j).setBackgroundColor('#f5f5f5');
    }
    for (let i = 1; i < synthTable.getNumRows(); i++) {
      synthTable.getRow(i).getCell(0).getChild(0).asParagraph().setAttributes({ [DocumentApp.Attribute.BOLD]: true });
    }
    for (let i = 0; i < synthTable.getNumRows(); i++) {
      const row = synthTable.getRow(i);
      for (let j = 0; j < row.getNumCells(); j++) {
        row.getCell(j).setPaddingTop(8).setPaddingBottom(8);
      }
    }
    
    body.appendParagraph(""); // Spacer
    
    // Barème d'Évaluation de la performance
    const legendHeading = body.appendParagraph("Barème d'Évaluation");
    legendHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    legendHeading.setAttributes({ [DocumentApp.Attribute.BOLD]: true });

    const legendTable = body.appendTable([
      ["Niveau", "Intitulé", "Définition"],
      ["5 - PE", "Performance Exceptionnelle", "Dépasse constamment les attentes et constitue une référence."],
      ["4 - PS", "Performance Supérieure", "Dépasse régulièrement les attentes dans les domaines clés."],
      ["3 - CA", "Performance Conforme aux Attentes", "Atteint les objectifs et les standards de performance attendus pour le poste."],
      ["2 - PA", "Performance à Améliorer", "N'atteint pas toujours les objectifs et nécessite une amélioration dans certains domaines."],
      ["1 - PI", "Performance Insatisfaisante", "N'atteint pas les objectifs de manière significative et nécessite une amélioration immédiate."],
      ["N/A", "Non Applicable", "Non applicable au poste ou non évaluable sur la période."]
    ]);
    const legHeader = legendTable.getRow(0);
    for (let j = 0; j < legHeader.getNumCells(); j++) {
      legHeader.getCell(j).getChild(0).asParagraph().setAttributes({ [DocumentApp.Attribute.BOLD]: true });
      legHeader.getCell(j).setBackgroundColor('#f5f5f5');
    }
    for (let i = 0; i < legendTable.getNumRows(); i++) {
      const lRow = legendTable.getRow(i);
      for (let j = 0; j < lRow.getNumCells(); j++) {
        lRow.getCell(j).setPaddingTop(4).setPaddingBottom(4);
      }
    }

    body.appendParagraph(""); // Spacer

    // Détail des Compétences
    const compHeading = body.appendParagraph("Détail des Compétences");
    compHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    compHeading.setAttributes({ [DocumentApp.Attribute.BOLD]: true });
    
    const compData = [
      ["Compétence", "Auto-évaluation (Employé)", "Évaluation(s) Secondaire(s)", "Évaluation Principale"]
    ];
    
    // Helper to safely parse JSON
    const safeParse = (str) => {
      try { return JSON.parse(str || '[]'); } catch (e) { return []; }
    };
    
    const empFonda = safeParse(data[10]);
    const empSpec = safeParse(data[11]);
    const mgrFonda = safeParse(data[21]);
    const mgrSpec = safeParse(data[22]);
    const secFonda = safeParse(data[34]);
    const secSpec = safeParse(data[35]);

    const formatAnswerCell = (qObj) => {
      if (!qObj) return '-';
      const ans = qObj.answer || '-';
      const comment = qObj.empComment || qObj.mgrComment || qObj.comment || '';
      if (comment) {
        return ans + "\nCommentaire : " + comment;
      }
      return ans;
    };

    const formatSecondaryCell = (secEntries, idx) => {
      if (!secEntries || secEntries.length === 0) return '-';
      const parts = [];
      secEntries.forEach(entry => {
        const resp = entry.responses && entry.responses[idx];
        if (resp) {
          const ans = resp.answer || '-';
          const comm = resp.comment || resp.mgrComment || resp.empComment || '';
          parts.push((entry.evaluatorName || 'Secondaire') + " : " + ans + (comm ? "\n(" + comm + ")" : ""));
        }
      });
      return parts.length > 0 ? parts.join("\n\n") : '-';
    };
    
    if (empFonda.length > 0 || mgrFonda.length > 0) {
      compData.push(["Fondamentales", "", "", ""]);
      const count = Math.max(empFonda.length, mgrFonda.length);
      for (let i = 0; i < count; i++) {
        const empQ = empFonda[i] || {};
        const mgrQ = mgrFonda[i] || {};
        const qText = empQ.question || mgrQ.question || `Question ${i+1}`;
        compData.push([qText, formatAnswerCell(empQ), formatSecondaryCell(secFonda, i), formatAnswerCell(mgrQ)]);
      }
    }
    
    if (empSpec.length > 0 || mgrSpec.length > 0) {
      compData.push(["Spécifiques (" + (data[9] || 'Profil') + ")", "", "", ""]);
      const count = Math.max(empSpec.length, mgrSpec.length);
      for (let i = 0; i < count; i++) {
        const empQ = empSpec[i] || {};
        const mgrQ = mgrSpec[i] || {};
        const qText = empQ.question || mgrQ.question || `Question ${i+1}`;
        compData.push([qText, formatAnswerCell(empQ), formatSecondaryCell(secSpec, i), formatAnswerCell(mgrQ)]);
      }
    }
    
    const compTable = body.appendTable(compData);
    
    // Format Comp Table
    const compHeader = compTable.getRow(0);
    for (let j = 0; j < compHeader.getNumCells(); j++) {
      compHeader.getCell(j).getChild(0).asParagraph().setAttributes({ [DocumentApp.Attribute.BOLD]: true });
      compHeader.getCell(j).setBackgroundColor('#f5f5f5');
    }
    for (let i = 1; i < compTable.getNumRows(); i++) {
      const row = compTable.getRow(i);
      const text = row.getCell(0).getText();
      if (text === "Fondamentales" || text.startsWith("Spécifiques")) {
        row.getCell(0).getChild(0).asParagraph().setAttributes({ [DocumentApp.Attribute.BOLD]: true });
        row.getCell(0).setBackgroundColor('#fafafa');
        row.getCell(1).setBackgroundColor('#fafafa');
        row.getCell(2).setBackgroundColor('#fafafa');
        row.getCell(3).setBackgroundColor('#fafafa');
      }
      for (let j = 0; j < row.getNumCells(); j++) {
        row.getCell(j).setPaddingTop(8).setPaddingBottom(8);
      }
    }
    
    doc.saveAndClose();
    
    const pdfBlob = doc.getAs('application/pdf');
    pdfBlob.setName(`${docName}.pdf`);
    
    let file;
    try {
      const folderId = '10y8QMWCUlWL1frurah6lXg2d8j1ukGlj';
      const folder = DriveApp.getFolderById(folderId);
      file = folder.createFile(pdfBlob);
    } catch(folderErr) {
      console.error("Erreur accès dossier Drive, fallback racine : " + folderErr.message);
      file = DriveApp.createFile(pdfBlob);
    }
    
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    
    return file.getUrl();
}

function saveAttachedManagerEvaluationDraft(token, rowId, formData, managerNameFallback) {
  let managerName = managerNameFallback || '';
  if (token && typeof token === 'string' && token.length > 10) {
    try {
      const sessionUser = verifySession(token);
      managerName = sessionUser.fullName || sessionUser.username;
    } catch(e) {}
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error("Feuille Evaluations introuvable.");
    
    const row = parseInt(rowId);
    const maxRequiredCol = 37;
    if (sheet.getMaxColumns() < maxRequiredCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxRequiredCol - sheet.getMaxColumns());
    }
    
    const existingMgr = sheet.getRange(row, 21).getValue();
    let updatedMgrName = managerName || existingMgr || 'Manager';
    if (existingMgr && managerName && existingMgr !== managerName && !String(existingMgr).includes(managerName)) {
      updatedMgrName = existingMgr + ", " + managerName;
    }
    
    sheet.getRange(row, 21).setValue(updatedMgrName);
    sheet.getRange(row, 22).setValue(JSON.stringify(formData.mgrFondamentales || []));
    sheet.getRange(row, 23).setValue(JSON.stringify(formData.mgrSpecifiques || []));
    sheet.getRange(row, 24).setValue(formData.mgrStrengths || '');
    sheet.getRange(row, 25).setValue(formData.mgrImprovements || '');
    sheet.getRange(row, 26).setValue(formData.mgrGlobalRating || '');
    sheet.getRange(row, 27).setValue(formData.mgrTraining || '');
    sheet.getRange(row, 28).setValue(formData.mgrSmartGoals || '');
    sheet.getRange(row, 29).setValue(formData.mgrComments || '');

    return { success: true };
  } catch (e) {
    throw new Error("Erreur lors de la sauvegarde du brouillon manager: " + e.message);
  }
}

function submitAttachedManagerEvaluation(token, rowId, formData, managerNameFallback) {
  let managerName = managerNameFallback || '';
  if (token && typeof token === 'string' && token.length > 10) {
    try {
      const sessionUser = verifySession(token);
      managerName = sessionUser.fullName || sessionUser.username;
    } catch(e) {}
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error("Feuille Evaluations introuvable.");
    
    const row = parseInt(rowId);
    const maxRequiredCol = 37;
    if (sheet.getMaxColumns() < maxRequiredCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxRequiredCol - sheet.getMaxColumns());
    }
    
    const existingMgr = sheet.getRange(row, 21).getValue();
    let updatedMgrName = managerName || existingMgr || 'Manager';
    if (existingMgr && managerName && existingMgr !== managerName && !String(existingMgr).includes(managerName)) {
      updatedMgrName = existingMgr + ", " + managerName;
    }
    
    sheet.getRange(row, 2).setValue('Complétée');
    sheet.getRange(row, 21).setValue(updatedMgrName);
    sheet.getRange(row, 22).setValue(JSON.stringify(formData.mgrFondamentales || []));
    sheet.getRange(row, 23).setValue(JSON.stringify(formData.mgrSpecifiques || []));
    sheet.getRange(row, 24).setValue(formData.mgrStrengths || '');
    sheet.getRange(row, 25).setValue(formData.mgrImprovements || '');
    sheet.getRange(row, 26).setValue(formData.mgrGlobalRating || '');
    sheet.getRange(row, 27).setValue(formData.mgrTraining || '');
    sheet.getRange(row, 28).setValue(formData.mgrSmartGoals || '');
    sheet.getRange(row, 29).setValue(formData.mgrComments || '');

    const pdfUrl = generateEvaluationPDF(row);
    if(pdfUrl) {
       sheet.getRange(row, 31).setValue(pdfUrl);
    }

    return { success: true };
  } catch (e) {
    throw new Error("Erreur lors de l'attachement de l'évaluation: " + e.message);
  }
}

function getNativeEvaluations(token) {
  const sessionUser = verifySession(token);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) return [];
    
    const lastCol = Math.max(sheet.getLastColumn() || 1, 37);
    const lastRow = sheet.getLastRow() || 1;
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const evals = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      let status = row[1] ? row[1].toString() : '';
      if (status.startsWith('Initi')) status = 'Initiée';
      else if (status.startsWith('Attente Évaluateur Secondaire') || status.startsWith('Attente \u00C9valuateur Secondaire')) status = 'Attente Évaluateur Secondaire';
      else if (status.startsWith('Attente Évaluateur Principal') || status.startsWith('Attente \u00C9valuateur Principal')) status = 'Attente Évaluateur Principal';
      else if (status.startsWith('Auto-')) status = 'Auto-évaluée';
      else if (status.startsWith('Compl')) status = 'Complétée';
      else if (status.startsWith('Pr')) status = 'Pré-évaluée';

      evals.push({
        rowId: i + 1,
        timestamp: row[0] instanceof Date ? row[0].toLocaleDateString('fr-FR') : row[0],
        status: status,
        period: row[2],
        employeeName: row[3],
        dept: row[4],
        entryDate: row[5] instanceof Date ? row[5].toLocaleDateString('fr-FR') : row[5],
        evaluatorName: row[6],
        evaluatorJob: row[7],
        interviewDate: row[8] instanceof Date ? row[8].toLocaleDateString('fr-FR') : row[8],
        profile: row[9],
        fondamentales: row[10],
        specifiques: row[11],
        pastGoals: row[12],
        strengths: row[13],
        improvements: row[14],
        globalRating: row[15],
        training: row[16],
        smartGoals: row[17],
        aspirations: row[18],
        empComments: row[19],
        mgrName: row[20],
        mgrFondamentales: row[21],
        mgrSpecifiques: row[22],
        mgrStrengths: row[23],
        mgrImprovements: row[24],
        mgrGlobalRating: row[25],
        mgrTraining: row[26],
        mgrSmartGoals: row[27],
        mgrComments: row[28],
        employeeEmail: row[29],
        pdfUrl: row[30] || '',
        principalEvaluatorEmail: row[31] || '',
        principalEvaluatorName: row[32] || row[20] || row[6] || '',
        secondaryEvaluators: row[33] || '[]',
        secondaryFondamentales: row[34] || '[]',
        secondarySpecifiques: row[35] || '[]',
        secondarySynthese: row[36] || '[]'
      });
    }
    return JSON.stringify(evals.reverse());
  } catch (e) {
    return JSON.stringify([{error: e.message}]); 
  }
}

function checkIfManager(username) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Utilisateurs');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === username) {
          const role = data[i][5] || 'Utilisateur';
          return { isManager: (role === 'Admin' || role === 'Manager'), isAdmin: (role === 'Admin') };
        }
      }
    }
  } catch (e) {
    console.error(e);
  }
  return { isManager: false, isAdmin: false }; 
}

// =================================================================
// 6. ADMINISTRATION SYSTEM
// =================================================================

function getEmployeeListForDropdown(token) {
  const sessionUser = verifySession(token);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Utilisateurs');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) list.push({ email: data[i][0], fullName: data[i][2] || data[i][0] });
    }
    return list;
  } catch (e) {
    console.error("Error getEmployeeListForDropdown: " + e);
    return [];
  }
}

function getAllUsers(token) {
  const sessionUser = verifySession(token);
  const adminUsername = sessionUser.username;

  try {
    const adminCheck = checkIfManager(adminUsername);
    if (!adminCheck.isAdmin) throw new Error("Accès refusé. Réservé aux administrateurs.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Utilisateurs');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const users = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      users.push({
        rowId: i + 1,
        email: data[i][0],
        fullName: data[i][2],
        company: data[i][3],
        role: data[i][5] || 'Utilisateur',
        allowedProjectsStr: data[i][6] || ''
      });
    }
    return users;
  } catch (e) {
    throw new Error("Erreur getAllUsers: " + e.message);
  }
}

function getAllProjectsForAdmin(token) {
  const sessionUser = verifySession(token);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const projectsSheet = ss.getSheetByName("Projects Odoo");
  const allProjects = new Set();
  if (projectsSheet) {
    const projectsData = projectsSheet.getRange(2, 2, projectsSheet.getLastRow() - 1, 1).getValues();
    projectsData.forEach(row => {
      if (row[0]) allProjects.add(row[0]);
    });
  }
  return Array.from(allProjects).sort();
}

function debugGetEvals() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Evaluations');
  if (!sheet) return "No sheet";
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow || 1, lastCol || 1).getValues();
  
  let result = `lastRow: ${lastRow}, lastCol: ${lastCol}\n`;
  for(let i=0; i<data.length; i++) {
    result += `Row ${i+1}: ColA='${data[i][0]}' type=${typeof data[i][0]}\n`;
  }
  
  const evals = getNativeEvaluations();
  result += `\ngetNativeEvaluations() returned: ${evals}`;
  
  return result;
}

function updateUserRoleAndPassword(token, userEmail, newRole, newPassword, newCompany, allowedProjectsStr) {
  const sessionUser = verifySession(token);
  const adminUsername = sessionUser.username;

  try {
    const adminCheck = checkIfManager(adminUsername);
    if (!adminCheck.isAdmin) throw new Error("Accès refusé.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Utilisateurs');
    if (!sheet) throw new Error("Feuille Utilisateurs introuvable.");

    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        if (newRole) sheet.getRange(i + 1, 6).setValue(newRole); // Col F
        if (newPassword) sheet.getRange(i + 1, 2).setValue(newPassword); // Col B
        if (newCompany !== undefined && newCompany !== null) sheet.getRange(i + 1, 4).setValue(newCompany); // Col D
        if (newRole === 'ConsultantExterne' && allowedProjectsStr !== undefined) {
           sheet.getRange(i + 1, 7).setValue(allowedProjectsStr); // Col G
        } else if (newRole && newRole !== 'ConsultantExterne') {
           sheet.getRange(i + 1, 7).setValue(''); // Clear if no longer a consultant
        }
        found = true;
        break;
      }
    }
    if (!found) throw new Error("Utilisateur introuvable.");
    return { success: true };
  } catch (e) {
    throw new Error("Erreur de mise à jour: " + e.message);
  }
}

function triggerOdooSync(token) {
  const sessionUser = verifySession(token);
  const adminUsername = sessionUser.username;

  try {
    const adminCheck = checkIfManager(adminUsername);
    if (!adminCheck.isAdmin) throw new Error("Accès refusé.");
    syncOdooEmployees(); // reuse existing function
    syncOdooTasks();
    return { success: true };
  } catch (e) {
    throw new Error("Erreur sync: " + e.message);
  }
}

function saveAttachedManagerEvaluationDraft(token, rowId, formData) {
  const sessionUser = verifySession(token);
  const managerName = sessionUser.fullName || sessionUser.username;

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error('Feuille Evaluations introuvable.');
    
    const row = parseInt(rowId);
    
    const maxRequiredCol = 31;
    if (sheet.getMaxColumns() < maxRequiredCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxRequiredCol - sheet.getMaxColumns());
    }
    
    const existingMgr = sheet.getRange(row, 21).getValue();
    let updatedMgrName = managerName;
    if (existingMgr && existingMgr !== managerName && !String(existingMgr).includes(managerName)) {
      updatedMgrName = existingMgr + ', ' + managerName;
    }
    
    const currentStatus = sheet.getRange(row, 2).getValue();
    if (currentStatus === 'Auto-\u00E9valu\u00E9e') {
      sheet.getRange(row, 2).setValue('Brouillon Manager');
    }
    
    sheet.getRange(row, 21).setValue(updatedMgrName);
    sheet.getRange(row, 22).setValue(JSON.stringify(formData.mgrFondamentales || []));
    sheet.getRange(row, 23).setValue(JSON.stringify(formData.mgrSpecifiques || []));
    sheet.getRange(row, 24).setValue(formData.mgrStrengths || '');
    sheet.getRange(row, 25).setValue(formData.mgrImprovements || '');
    sheet.getRange(row, 26).setValue(formData.mgrGlobalRating || '');
    sheet.getRange(row, 27).setValue(formData.mgrTraining || '');
    sheet.getRange(row, 28).setValue(formData.mgrSmartGoals || '');
    sheet.getRange(row, 29).setValue(formData.mgrComments || '');

    return { success: true };
  } catch (e) {
    throw new Error('Erreur lors de la sauvegarde du brouillon manager: ' + e.message);
  }
}


function savePreEvaluation(token, rowId, formData) {
  const sessionUser = verifySession(token);
  const managerName = sessionUser.fullName || sessionUser.username;

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Evaluations');
    if (!sheet) throw new Error('Feuille Evaluations introuvable.');
    
    const row = parseInt(rowId);
    
    const maxRequiredCol = 31;
    if (sheet.getMaxColumns() < maxRequiredCol) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxRequiredCol - sheet.getMaxColumns());
    }
    
    const existingMgr = sheet.getRange(row, 21).getValue();
    let updatedMgrName = managerName;
    if (existingMgr && existingMgr !== managerName && !String(existingMgr).includes(managerName)) {
      updatedMgrName = existingMgr + ', ' + managerName;
    }
    
    sheet.getRange(row, 2).setValue('Pré-évaluée');
    
    sheet.getRange(row, 21).setValue(updatedMgrName);
    sheet.getRange(row, 22).setValue(JSON.stringify(formData.mgrFondamentales || []));
    sheet.getRange(row, 23).setValue(JSON.stringify(formData.mgrSpecifiques || []));
    sheet.getRange(row, 24).setValue(formData.mgrStrengths || '');
    sheet.getRange(row, 25).setValue(formData.mgrImprovements || '');
    sheet.getRange(row, 26).setValue(formData.mgrGlobalRating || '');
    sheet.getRange(row, 27).setValue(formData.mgrTraining || '');
    sheet.getRange(row, 28).setValue(formData.mgrSmartGoals || '');
    sheet.getRange(row, 29).setValue(formData.mgrComments || '');

    return { success: true };
  } catch (e) {
    throw new Error('Erreur lors de la sauvegarde de la pré-évaluation: ' + e.message);
  }
}

// ==========================================
// TICKET SUPPORT SYSTEM
// ==========================================

function getTicketsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Tickets');
  if (!sheet) {
    sheet = ss.insertSheet('Tickets');
    // ID, Date, UserEmail, Subject, Priority, Description, Status, IT Comment
    sheet.appendRow(['ID', 'Date', 'Email', 'Sujet', 'Priorité', 'Description', 'Statut', 'Commentaire IT']);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function submitTicket(token, ticketData) {
  const sessionUser = verifySession(token);
  
  try {
    const sheet = getTicketsSheet();
    const ticketId = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    sheet.appendRow([
      ticketId,
      dateStr,
      sessionUser.username,
      ticketData.subject,
      ticketData.priority,
      ticketData.description,
      'Ouvert',
      ''
    ]);
    
    return { success: true };
  } catch(e) {
    throw new Error("Erreur lors de la soumission du ticket: " + e.message);
  }
}

function getUserTickets(token) {
  const sessionUser = verifySession(token);
  
  try {
    const sheet = getTicketsSheet();
    const data = sheet.getDataRange().getValues();
    const tickets = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === sessionUser.username) {
        tickets.push({
          id: data[i][0],
          date: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(data[i][1] || ''),
          userEmail: data[i][2],
          subject: data[i][3],
          priority: data[i][4],
          description: data[i][5],
          status: data[i][6],
          comment: data[i][7]
        });
      }
    }
    return tickets.reverse(); // Newest first
  } catch(e) {
    throw new Error("Erreur de récupération des tickets: " + e.message);
  }
}

function getAllTickets(token) {
  const sessionUser = verifySession(token);
  if (sessionUser.role !== 'Admin' && sessionUser.role !== 'IT') {
    throw new Error("Accès refusé.");
  }
  
  try {
    const sheet = getTicketsSheet();
    const data = sheet.getDataRange().getValues();
    const tickets = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        tickets.push({
          id: data[i][0],
          date: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(data[i][1] || ''),
          userEmail: data[i][2],
          subject: data[i][3],
          priority: data[i][4],
          description: data[i][5],
          status: data[i][6],
          comment: data[i][7]
        });
      }
    }
    return tickets.reverse();
  } catch(e) {
    throw new Error("Erreur de récupération de tous les tickets: " + e.message);
  }
}

function updateTicket(token, ticketId, newStatus, newComment) {
  const sessionUser = verifySession(token);
  if (sessionUser.role !== 'Admin' && sessionUser.role !== 'IT') {
    throw new Error("Accès refusé.");
  }
  
  try {
    const sheet = getTicketsSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === ticketId) {
        sheet.getRange(i + 1, 7).setValue(newStatus);
        sheet.getRange(i + 1, 8).setValue(newComment);
        return { success: true };
      }
    }
    throw new Error("Ticket introuvable.");
  } catch(e) {
    throw new Error("Erreur lors de la mise à jour: " + e.message);
  }
}

/**
 * Configuration par défaut des groupes d'administration par pays pour l'accès aux réponses du formulaire.
 */
const DEFAULT_COUNTRY_ADMIN_GROUPS = {
  'Mali': [
    'amdiallo@kanagaconsulting.com',
    'rdsacko@kanagaconsulting.com',
    'sdolo@kanagaconsulting.com'
  ],
  'Burkina Faso': [
    'lhounkponou@kanagaconsulting.com',
    'rdsacko@kanagaconsulting.com',
    'skaloga@kanagaconsulting.com'
  ],
  'République démocratique du Congo': [
    'rdsacko@kanagaconsulting.com',
    'skaloga@kanagaconsulting.com',
    'sdolo@kanagaconsulting.com'
  ],
  "Côte d'Ivoire": [
    'gsyabre@kanagaconsulting.com',
    'rdsacko@kanagaconsulting.com',
    'skaloga@kanagaconsulting.com',
    'sdolo@kanagaconsulting.com'
  ],
  'Guinée': [
    'gcamara@kanagaconsulting.com',
    'itounkara@kanagaconsulting.com',
    'sdolo@kanagaconsulting.com'
  ]
};

/**
 * Récupère la configuration des groupes d'administration par pays depuis ScriptProperties (ou initialise les valeurs par défaut).
 */
function getCountryAdminGroupsConfig(token) {
  if (token) verifySession(token);
  const props = PropertiesService.getScriptProperties();
  const storedStr = props.getProperty('COUNTRY_ADMIN_GROUPS_V4');
  if (storedStr) {
    try {
      const parsed = JSON.parse(storedStr);
      let ciKey = Object.keys(parsed).find(k => k.toLowerCase().includes("cote d'ivoire") || k.toLowerCase().includes("côte d'ivoire"));
      if (ciKey) {
        const lowerList = (parsed[ciKey] || []).map(e => e.toLowerCase());
        if (!lowerList.includes('gsyabre@kanagaconsulting.com')) {
          parsed[ciKey].push('gsyabre@kanagaconsulting.com');
        }
      } else {
        parsed["Côte d'Ivoire"] = ['gsyabre@kanagaconsulting.com', 'rdsacko@kanagaconsulting.com', 'skaloga@kanagaconsulting.com', 'sdolo@kanagaconsulting.com'];
      }
      return parsed;
    } catch(e) {}
  }
  
  props.setProperty('COUNTRY_ADMIN_GROUPS_V4', JSON.stringify(DEFAULT_COUNTRY_ADMIN_GROUPS));
  return DEFAULT_COUNTRY_ADMIN_GROUPS;
}

/**
 * Enregistre la configuration mise à jour des groupes d'administration par pays depuis le panneau Administration.
 */
function saveCountryAdminGroupsConfig(token, newGroupsData) {
  const sessionUser = verifySession(token);
  if (sessionUser.role !== 'Admin') {
    throw new Error("Accès refusé. Seuls les administrateurs globaux peuvent modifier la configuration des groupes.");
  }
  
  if (!newGroupsData || typeof newGroupsData !== 'object') {
    throw new Error("Données de groupes invalides.");
  }

  PropertiesService.getScriptProperties().setProperty('COUNTRY_ADMIN_GROUPS_V4', JSON.stringify(newGroupsData));
  return { success: true };
}

/**
 * Vérifie si l'utilisateur connecté a le droit de voir l'onglet "Réponses Formulaire" et quels bureaux il peut voir.
 */
function checkFormResponsesAccess(token) {
  const sessionUser = verifySession(token);
  const userEmail = String(sessionUser.username || '').toLowerCase().trim();
  const isAdmin = (sessionUser.role === 'Admin');

  const groups = getCountryAdminGroupsConfig();
  const allowedCountries = new Set();

  if (isAdmin) {
    return { canAccess: true, isAdmin: true };
  }

  for (const [countryKey, allowedEmails] of Object.entries(groups)) {
    const lowerAllowed = (allowedEmails || []).map(e => String(e).toLowerCase().trim());
    if (lowerAllowed.includes(userEmail)) {
      allowedCountries.add(countryKey);
    }
  }

  // Vérifier également si l'utilisateur possède une société assignée
  const userSocieties = sessionUser.societies || [];
  const activeSociety = sessionUser.activeSociety || '';

  return {
    canAccess: allowedCountries.size > 0 || userSocieties.length > 0 || activeSociety !== '',
    isAdmin: false,
    allowedCountries: Array.from(allowedCountries)
  };
}

/**
 * Récupère les réponses du formulaire "Réponses au formulaire 1" avec filtrage dynamique selon les groupes d'administration par pays.
 */
function normalizeStr(str) {
  if (!str) return '';
  return String(str)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\\u0300-\\u036f]/g, "")
    .trim();
}

function matchCountry(bureauStr, countryKey) {
  const normBureau = normalizeStr(bureauStr);
  const normCountry = normalizeStr(countryKey);

  if (!normBureau || !normCountry) return false;
  if (normBureau.includes(normCountry) || normCountry.includes(normBureau)) {
    return true;
  }

  if (normCountry.includes('cote d\'ivoire') || normCountry.includes('ci') || normCountry.includes('abidjan')) {
    if (normBureau.includes('cote d\'ivoire') || normBureau.includes('ci') || normBureau.includes('abidjan') || normBureau.includes('ivory coast')) return true;
  }
  if (normCountry.includes('congo') || normCountry.includes('rdc')) {
    if (normBureau.includes('congo') || normBureau.includes('rdc')) return true;
  }
  if (normCountry.includes('guinee')) {
    if (normBureau.includes('guinee') || normBureau.includes('guinea')) return true;
  }
  if (normCountry.includes('mali')) {
    if (normBureau.includes('mali')) return true;
  }
  if (normCountry.includes('burkina')) {
    if (normBureau.includes('burkina') || normBureau.includes('faso')) return true;
  }

  return false;
}

function getFormResponses(token, selectedBureau) {
  const sessionUser = verifySession(token);
  const userEmail = String(sessionUser.username || '').toLowerCase().trim();
  const isAdmin = (sessionUser.role === 'Admin');
  const userSocieties = sessionUser.societies || [];
  const activeSociety = sessionUser.activeSociety || '';

  const countryAdminGroups = getCountryAdminGroupsConfig();

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Réponses au formulaire 1');
    if (!sheet) {
      throw new Error("La feuille 'Réponses au formulaire 1' est introuvable dans le classeur Google Sheets.");
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      return {
        headers: [],
        responses: [],
        bureaux: [],
        userAllowedBureaux: [],
        afColIndex: 31,
        userBureau: activeSociety,
        isAdmin: isAdmin,
        hasAccess: true
      };
    }

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0].map(h => String(h || '').trim());

    let afIndex = -1;
    for (let c = 0; c < headers.length; c++) {
      const hLower = headers[c].toLowerCase();
      if (hLower.indexOf('bureau kanaga') !== -1 || hLower.indexOf('à quel bureau') !== -1) {
        afIndex = c;
        break;
      }
    }
    if (afIndex === -1 && headers.length >= 32) {
      afIndex = 31;
    }

    const allBureauxSet = new Set();
    if (afIndex !== -1) {
      for (let i = 1; i < data.length; i++) {
        const val = String(data[i][afIndex] || '').trim();
        if (val) allBureauxSet.add(val);
      }
    }
    const bureauxList = Array.from(allBureauxSet).sort();

    function isUserAuthorizedForBureau(bureauStr) {
      if (isAdmin) return true;
      if (!bureauStr) return false;

      let isRestrictedCountry = false;

      for (const [countryKey, allowedEmails] of Object.entries(countryAdminGroups)) {
        if (matchCountry(bureauStr, countryKey)) {
          isRestrictedCountry = true;
          const lowerAllowed = (allowedEmails || []).map(e => String(e).toLowerCase().trim());
          if (lowerAllowed.includes(userEmail)) {
            return true;
          }
        }
      }

      if (isRestrictedCountry) {
        return false;
      }

      for (let s of userSocieties) {
        if (!s) continue;
        if (matchCountry(bureauStr, s)) return true;
      }

      if (activeSociety && matchCountry(bureauStr, activeSociety)) {
        return true;
      }

      return false;
    }

    const userAllowedBureaux = bureauxList.filter(b => isUserAuthorizedForBureau(b));
    const showBureauSelector = isAdmin || (userAllowedBureaux.length > 1);

    function matchesSelectedFilter(rowBureauStr) {
      if (!isUserAuthorizedForBureau(rowBureauStr)) return false;

      if (!selectedBureau || selectedBureau === 'ALL') return true;
      return rowBureauStr.toLowerCase() === selectedBureau.toLowerCase();
    }

    const responses = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowBureau = (afIndex !== -1 && afIndex < row.length) ? String(row[afIndex] || '').trim() : '';

      if (matchesSelectedFilter(rowBureau)) {
        const formattedRow = row.map(val => {
          if (val instanceof Date) {
            return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          }
          return String(val !== null && val !== undefined ? val : '');
        });
        responses.push({
          rowIndex: i + 1,
          bureau: rowBureau,
          values: formattedRow
        });
      }
    }

    return {
      headers: headers,
      responses: responses.reverse(),
      bureaux: bureauxList,
      userAllowedBureaux: userAllowedBureaux,
      afColIndex: afIndex,
      userBureau: activeSociety,
      isAdmin: isAdmin,
      showBureauSelector: showBureauSelector,
      hasAccess: isAdmin || (userAllowedBureaux.length > 0)
    };
  } catch (e) {
    console.error("Erreur getFormResponses: " + e.message);
    throw new Error("Erreur de récupération des réponses du formulaire: " + e.message);
  }
}
