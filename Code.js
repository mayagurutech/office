// ==========================================
// 1. WEB APP DEPLOYMENT (Templates)
// ==========================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Office OS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 2. AI NOTE GENERATION (GEMINI)
// ==========================================
function generateAINote(roughNotes) {
  const apiKey = 'AIzaSyAH2aKEJX1iYHKxK_9zUSObAhbvZfJSdmc'; 
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
  const prompt = "Rewrite these rough debtor follow-up notes into a brief, professional summary for a firm's internal log. Do not use greetings or pleasantries. Keep it to 1-2 concise sentences. Rough notes: " + roughNotes;
  const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates.length > 0) return json.candidates[0].content.parts[0].text.trim();
    else throw new Error("Invalid AI response");
  } catch (e) { throw new Error("Failed to generate AI note: " + e.message); }
}

// ==========================================
// 3. ONE-TIME MIGRATION SCRIPTS
// ==========================================
const supabaseUrlRest = 'https://xeebucvvaxyglwjkkfvr.supabase.co/rest/v1';
const supabaseKeyAuth = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhlZWJ1Y3Z2YXh5Z2x3amtrZnZyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ3Njc2NzUsImV4cCI6MjA5MDM0MzY3NX0.jrZXmI6gx-gyhczCC2cBSSaVh8KTfjCjhYapNEaDBEM';

function migrateClientsToSupabase() {
  const auditData = SpreadsheetApp.openById('1BVOTgOPsvGnpefO1zWeDeGwa-aLu3_qdm_kTE4JsUGY').getSheetByName('Clients').getDataRange().getValues();
  const debtData = SpreadsheetApp.openById('16jRkyPkZJygY6QsTU3L8urXQVZyGG-FvY9DHyuJv1-8').getSheetByName('Master').getDataRange().getValues();
  let uniqueClients = {};
  for (let i = 1; i < auditData.length; i++) { if (auditData[i][0]) uniqueClients[auditData[i][0].toString().trim()] = true; }
  for (let i = 1; i < debtData.length; i++) { if (debtData[i][0]) uniqueClients[debtData[i][0].toString().trim()] = true; }
  let payload = []; for (let clientName in uniqueClients) payload.push({ display_name: clientName, entity_type: "Migrated from Sheets" });
  if (payload.length > 0) UrlFetchApp.fetch(`${supabaseUrlRest}/clients`, { 'method' : 'post', 'contentType': 'application/json', 'headers': { 'apikey': supabaseKeyAuth, 'Authorization': 'Bearer ' + supabaseKeyAuth, 'Prefer': 'return=minimal' }, 'payload' : JSON.stringify(payload) });
}

// ==========================================
// 4. FIX SCRIPT: AUDIT USERS & HISTORY
// ==========================================
function fixAuditUsersAndHistory() {
  const optionsGet = { 'method': 'get', 'headers': { 'apikey': supabaseKeyAuth, 'Authorization': 'Bearer ' + supabaseKeyAuth } };
  const optionsPost = { 'method': 'post', 'contentType': 'application/json', 'headers': { 'apikey': supabaseKeyAuth, 'Authorization': 'Bearer ' + supabaseKeyAuth, 'Prefer': 'return=minimal' } };
  const optionsDelete = { 'method': 'delete', 'headers': { 'apikey': supabaseKeyAuth, 'Authorization': 'Bearer ' + supabaseKeyAuth } };

  // 1. Wipe existing task logs to prevent duplicates
  UrlFetchApp.fetch(supabaseUrlRest + '/task_logs?log_id=not.is.null', optionsDelete);
  
  // 2. Fetch Existing Supabase Users
  let existingUsers = JSON.parse(UrlFetchApp.fetch(supabaseUrlRest + '/users?select=*', optionsGet).getContentText());
  let existingNames = existingUsers.map(u => u.name.trim().toLowerCase());
  let existingPins = existingUsers.map(u => u.pin.toString());

  // 3. Fetch missing users from the Audit Sheet
  const auditSheetId = '1BVOTgOPsvGnpefO1zWeDeGwa-aLu3_qdm_kTE4JsUGY';
  const auditUsersData = SpreadsheetApp.openById(auditSheetId).getSheetByName('Users').getDataRange().getValues();
  
  let newUsersPayload = [];
  for (let i = 1; i < auditUsersData.length; i++) {
    let rawName = auditUsersData[i][0] ? auditUsersData[i][0].toString().trim() : null;
    let rawPin = auditUsersData[i][1] ? auditUsersData[i][1].toString().trim() : null;
    
    // If the name isn't in Supabase yet, add them
    if (rawName && !existingNames.includes(rawName.toLowerCase())) {
       let finalPin = rawPin;
       let finalName = rawName;
       
       // Handle Duplicate PINs
       if (existingPins.includes(rawPin)) {
         finalPin = Math.floor(1000 + Math.random() * 9000).toString(); // Random pin
         finalName = rawName + " (Duplicate PIN Auto-Fixed)"; // Flag for manual review
       }
       newUsersPayload.push({ name: finalName, pin: finalPin, role: 'Manager' });
       existingPins.push(finalPin); 
    }
  }

  // 4. Insert new missing users
  if (newUsersPayload.length > 0) {
    optionsPost.payload = JSON.stringify(newUsersPayload);
    UrlFetchApp.fetch(supabaseUrlRest + '/users', optionsPost);
    Logger.log("Migrated " + newUsersPayload.length + " missing users from Audit sheet.");
  }

  // 5. Re-fetch users to get their fresh UUIDs for mapping
  existingUsers = JSON.parse(UrlFetchApp.fetch(supabaseUrlRest + '/users?select=id,name', optionsGet).getContentText());
  const userMap = {}; 
  existingUsers.forEach(u => {
    let lowerName = u.name.trim().toLowerCase().replace(' (duplicate pin auto-fixed)', '');
    userMap[lowerName] = u.id;
    userMap[lowerName.split(' ')[0]] = u.id; // Map first names too
  });

  // 6. Fetch Clients & Tasks for mapping
  const clients = JSON.parse(UrlFetchApp.fetch(supabaseUrlRest + '/clients?select=client_id,display_name', optionsGet).getContentText());
  const tasks = JSON.parse(UrlFetchApp.fetch(supabaseUrlRest + '/tasks?select=task_id,client_id', optionsGet).getContentText());
  const clientMap = {}; clients.forEach(c => clientMap[c.display_name.trim().toLowerCase()] = c.client_id);
  const taskMap = {}; tasks.forEach(t => taskMap[t.client_id] = t.task_id);

  // 7. Fetch Work_Logs and migrate History
  const historySheet = SpreadsheetApp.openById(auditSheetId).getSheetByName('Work_Logs');
  const historyData = historySheet.getDataRange().getValues();
  let logsPayload = [];
  
  for (let i = 1; i < historyData.length; i++) {
    let clientName = historyData[i][1] ? historyData[i][1].toString().trim().toLowerCase() : null;
    let rawUserName = historyData[i][2] ? historyData[i][2].toString().trim().toLowerCase() : null;
    
    let clientId = clientMap[clientName];
    let taskId = clientId ? taskMap[clientId] : null;
    
    let userId = userMap[rawUserName];
    if (!userId && rawUserName) {
       const foundUser = existingUsers.find(u => u.name.toLowerCase().includes(rawUserName));
       if(foundUser) userId = foundUser.id;
    }

    if (taskId) {
      logsPayload.push({
        task_id: taskId,
        user_id: userId || null,
        timestamp: historyData[i][0] ? new Date(historyData[i][0]).toISOString() : new Date().toISOString(),
        bank_status: historyData[i][3] ? historyData[i][3].toString() : '',
        sales_status: historyData[i][4] ? historyData[i][4].toString() : '',
        ledger_status: historyData[i][5] ? historyData[i][5].toString() : '',
        others_status: historyData[i][6] ? historyData[i][6].toString() : '',
        note: historyData[i][7] ? historyData[i][7].toString() : 'Historical Update'
      });
    }
  }

  if (logsPayload.length > 0) {
    optionsPost.payload = JSON.stringify(logsPayload);
    UrlFetchApp.fetch(supabaseUrlRest + '/task_logs', optionsPost);
    Logger.log("Audit History rebuilt successfully. " + logsPayload.length + " rows saved.");
  }
}// force push  

// Sync push trigger
