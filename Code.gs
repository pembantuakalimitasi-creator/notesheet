/**
 * Google Apps Script - Notes & Todo Web App Backend
 * Supports both standalone Web App mode (google.script.run)
 * and GitHub Pages API mode (fetch POST).
 */

/**
 * Serves the web app UI for standalone mode
 */
function doGet(e) {
  if (e.parameter.api === 'true') {
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "API is active" }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({
        "Access-Control-Allow-Origin": "*"
      });
  }
  
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Catatan & Todo App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handle preflight CORS request
 */
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeaders({
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });
}

/**
 * Verify Google ID Token
 */
function verifyIdToken(idToken) {
  if (!idToken) throw new Error("Akses ditolak: Token tidak ditemukan");
  
  try {
    const url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + idToken;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const tokenInfo = JSON.parse(response.getContentText());
      return tokenInfo.email;
    }
  } catch (e) {
    // Fallback or error
  }
  throw new Error("Akses ditolak: Token tidak valid");
}

/**
 * Handles API POST requests (from GitHub Pages)
 */
function doPost(e) {
  var response;
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var params = requestData.params || {};
    var idToken = requestData.idToken;
    
    // Verify user and get email
    var userEmail = verifyIdToken(idToken);
    
    var data;
    if (action === 'getUserInfo') {
      data = { email: userEmail };
    } else if (action === 'getNotes') {
      data = getNotesForUser(userEmail);
    } else if (action === 'saveNote') {
      data = saveNoteForUser(userEmail, params.noteData);
    } else if (action === 'deleteNote') {
      data = deleteNoteForUser(userEmail, params.noteId);
    } else if (action === 'getTodos') {
      data = getTodosForUser(userEmail);
    } else if (action === 'addTodo') {
      data = addTodoForUser(userEmail, params.todoData);
    } else if (action === 'toggleTodo') {
      data = toggleTodoForUser(userEmail, params.todoId, params.completed);
    } else if (action === 'deleteTodo') {
      data = deleteTodoForUser(userEmail, params.todoId);
    } else if (action === 'getSpreadsheetUrl') {
      data = getSpreadsheetUrlForUser(userEmail);
    } else {
      throw new Error("Action tidak dikenal: " + action);
    }
    
    response = { success: true, data: data };
  } catch (err) {
    response = { success: false, message: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "Content-Type"
    });
}

/**
 * Returns active user info (for standalone mode)
 */
function getUserInfo() {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return {
    email: email || "Pengguna Google"
  };
}

/**
 * Helper to get or create the Spreadsheet database for a specific user
 */
function getDatabaseSpreadsheet(userEmail) {
  // Use Script Properties since user properties are shared in "Execute as Me"
  const scriptProperties = PropertiesService.getScriptProperties();
  const propKey = 'DB_' + userEmail.replace(/[^a-zA-Z0-9]/g, '_');
  let spreadsheetId = scriptProperties.getProperty(propKey);
  let ss;

  if (spreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      spreadsheetId = null;
    }
  }

  if (!spreadsheetId) {
    ss = SpreadsheetApp.create("Catatan & Todo Database - " + userEmail);
    
    let notesSheet = ss.getActiveSheet();
    notesSheet.setName("Notes");
    notesSheet.appendRow(["ID", "Title", "Content", "Category", "Color", "CreatedAt", "UpdatedAt"]);
    notesSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e2e8f0");
    
    let todosSheet = ss.insertSheet("Todos");
    todosSheet.appendRow(["ID", "Task", "Category", "Completed", "CreatedAt", "CompletedAt"]);
    todosSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#e2e8f0");
    
    scriptProperties.setProperty(propKey, ss.getId());
  }

  return ss;
}

/**
 * --- NOTES CRUD OPERATIONS (STANDALONE WRAPPERS) ---
 */
function getNotes() {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return getNotesForUser(email);
}

function saveNote(noteData) {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return saveNoteForUser(email, noteData);
}

function deleteNote(noteId) {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return deleteNoteForUser(email, noteId);
}

/**
 * --- TODOS CRUD OPERATIONS (STANDALONE WRAPPERS) ---
 */
function getTodos() {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return getTodosForUser(email);
}

function addTodo(todoData) {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return addTodoForUser(email, todoData);
}

function toggleTodo(todoId, completed) {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return toggleTodoForUser(email, todoId, completed);
}

function deleteTodo(todoId) {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return deleteTodoForUser(email, todoId);
}

function getSpreadsheetUrl() {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return getSpreadsheetUrlForUser(email);
}

/**
 * --- USER-SPECIFIC CRUD IMPLEMENTATIONS ---
 */

function getNotesForUser(userEmail) {
  const ss = getDatabaseSpreadsheet(userEmail);
  const sheet = ss.getSheetByName("Notes");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const notes = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const note = {};
    headers.forEach((header, index) => {
      note[header] = row[index];
    });
    notes.push(note);
  }
  
  return notes.sort((a, b) => new Date(b.UpdatedAt || b.CreatedAt) - new Date(a.UpdatedAt || a.CreatedAt));
}

function saveNoteForUser(userEmail, noteData) {
  const ss = getDatabaseSpreadsheet(userEmail);
  const sheet = ss.getSheetByName("Notes");
  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();
  
  let isNew = !noteData.ID;
  let id = noteData.ID || "note_" + Utilities.getUuid();
  
  if (isNew) {
    sheet.appendRow([
      id,
      noteData.Title || "Tanpa Judul",
      noteData.Content || "",
      noteData.Category || "Umum",
      noteData.Color || "#ffffff",
      now,
      now
    ]);
  } else {
    let foundIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === noteData.ID) {
        foundIndex = i + 1;
        break;
      }
    }
    
    if (foundIndex !== -1) {
      sheet.getRange(foundIndex, 2).setValue(noteData.Title || "Tanpa Judul");
      sheet.getRange(foundIndex, 3).setValue(noteData.Content || "");
      sheet.getRange(foundIndex, 4).setValue(noteData.Category || "Umum");
      sheet.getRange(foundIndex, 5).setValue(noteData.Color || "#ffffff");
      sheet.getRange(foundIndex, 7).setValue(now);
    } else {
      throw new Error("Catatan tidak ditemukan");
    }
  }
  
  return { success: true, message: isNew ? "Catatan berhasil ditambahkan!" : "Catatan berhasil diperbarui!" };
}

function deleteNoteForUser(userEmail, noteId) {
  const ss = getDatabaseSpreadsheet(userEmail);
  const sheet = ss.getSheetByName("Notes");
  const data = sheet.getDataRange().getValues();
  
  let foundIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === noteId) {
      foundIndex = i + 1;
      break;
    }
  }
  
  if (foundIndex !== -1) {
    sheet.deleteRow(foundIndex);
    return { success: true, message: "Catatan berhasil dihapus!" };
  } else {
    throw new Error("Catatan tidak ditemukan");
  }
}

function getTodosForUser(userEmail) {
  const ss = getDatabaseSpreadsheet(userEmail);
  const sheet = ss.getSheetByName("Todos");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const todos = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const todo = {};
    headers.forEach((header, index) => {
      if (header === "Completed") {
        todo[header] = row[index] === true || row[index] === "true";
      } else {
        todo[header] = row[index];
      }
    });
    todos.push(todo);
  }
  
  return todos.sort((a, b) => {
    if (a.Completed !== b.Completed) {
      return a.Completed ? 1 : -1;
    }
    return new Date(b.CreatedAt) - new Date(a.CreatedAt);
  });
}

function addTodoForUser(userEmail, todoData) {
  const ss = getDatabaseSpreadsheet(userEmail);
  const sheet = ss.getSheetByName("Todos");
  const now = new Date().toISOString();
  const id = "todo_" + Utilities.getUuid();
  
  sheet.appendRow([
    id,
    todoData.Task || "Tugas Tanpa Nama",
    todoData.Category || "Umum",
    false,
    now,
    ""
  ]);
  
  return { success: true, message: "Tugas berhasil ditambahkan!" };
}

function toggleTodoForUser(userEmail, todoId, completed) {
  const ss = getDatabaseSpreadsheet(userEmail);
  const sheet = ss.getSheetByName("Todos");
  const data = sheet.getDataRange().getValues();
  const now = completed ? new Date().toISOString() : "";
  
  let foundIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === todoId) {
      foundIndex = i + 1;
      break;
    }
  }
  
  if (foundIndex !== -1) {
    sheet.getRange(foundIndex, 4).setValue(completed);
    sheet.getRange(foundIndex, 6).setValue(now);
    return { success: true, message: completed ? "Tugas diselesaikan!" : "Tugas diaktifkan kembali!" };
  } else {
    throw new Error("Tugas tidak ditemukan");
  }
}

function deleteTodoForUser(userEmail, todoId) {
  const ss = getDatabaseSpreadsheet(userEmail);
  const sheet = ss.getSheetByName("Todos");
  const data = sheet.getDataRange().getValues();
  
  let foundIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === todoId) {
      foundIndex = i + 1;
      break;
    }
  }
  
  if (foundIndex !== -1) {
    sheet.deleteRow(foundIndex);
    return { success: true, message: "Tugas berhasil dihapus!" };
  } else {
    throw new Error("Tugas tidak ditemukan");
  }
}

function getSpreadsheetUrlForUser(userEmail) {
  try {
    const ss = getDatabaseSpreadsheet(userEmail);
    return ss.getUrl();
  } catch (e) {
    return null;
  }
}
