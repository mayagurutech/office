// ==========================================
// 1. WEB APP DEPLOYMENT (Templates)
// ==========================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Office OS')
      // DEFAULT stops other websites from embedding the app in a frame (clickjacking).
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ping() {
  return "pong";
}

// ==========================================
// 2. AI NOTE GENERATION (OpenRouter)
// ==========================================
// Uses the OpenRouter key stored in app_settings (setting_key = office_os_gemini_key).
// That row is hidden from staff browsers by RLS; only this script's secret key can read it.
// The web app is public, so the caller must send a valid staff Supabase access token.
const AI_NOTE_MAX_CHARS = 2000;
const AI_NOTE_LIMIT_PER_HOUR = 30;

function generateAINote(roughNotes, accessToken) {
  const staffId = requireStaff_(accessToken);

  roughNotes = String(roughNotes || '').trim();
  if (!roughNotes) throw new Error("Please type some rough notes first.");
  if (roughNotes.length > AI_NOTE_MAX_CHARS) throw new Error("Notes are too long for AI drafting (max " + AI_NOTE_MAX_CHARS + " characters).");

  // Per-staff hourly limit so a stolen session cannot run up the OpenRouter bill.
  const cache = CacheService.getScriptCache();
  const rateKey = 'ai_note_' + staffId + '_' + Math.floor(Date.now() / 3600000);
  const used = Number(cache.get(rateKey) || 0);
  if (used >= AI_NOTE_LIMIT_PER_HOUR) throw new Error("AI drafting limit reached for this hour. Please try again later.");
  cache.put(rateKey, String(used + 1), 3600);

  let apiKey = '';
  try {
    const settingsUrl = supabaseUrlRest + '/app_settings?setting_key=eq.office_os_gemini_key&select=setting_value';
    const settingsResponse = UrlFetchApp.fetch(settingsUrl, {
      'method': 'get',
      'headers': supabaseHeaders_(),
      'muteHttpExceptions': true
    });
    if (settingsResponse.getResponseCode() === 200) {
      const settingsJson = JSON.parse(settingsResponse.getContentText());
      if (settingsJson && settingsJson.length > 0 && settingsJson[0].setting_value) {
        apiKey = settingsJson[0].setting_value.trim();
      }
    }
  } catch (e) {
    Logger.log("Failed to fetch API key from Supabase: " + e.message);
  }

  if (!apiKey) {
    throw new Error("AI drafting is not configured: add an OpenRouter key as office_os_gemini_key in app_settings.");
  }

  const prompt = "Rewrite these rough debtor follow-up notes into a brief, professional summary for a firm's internal log. Do not use greetings or pleasantries. Keep it to 1-2 concise sentences. Rough notes: " + roughNotes;

  const url = 'https://openrouter.ai/api/v1/chat/completions';
  const payload = {
    "model": "google/gemini-2.5-flash",
    "messages": [{ "role": "user", "content": prompt }],
    "max_tokens": 150
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const resText = response.getContentText();
    const json = JSON.parse(resText);
    if (json.choices && json.choices.length > 0 && json.choices[0].message) {
      return json.choices[0].message.content.trim();
    } else {
      throw new Error("Invalid OpenRouter response: " + (json.error ? json.error.message : resText));
    }
  } catch (e) {
    throw new Error("Failed to generate AI note (OpenRouter): " + e.message);
  }
}

// ==========================================
// 3. ONE-TIME MIGRATION SCRIPTS
// ==========================================
const supabaseUrlRest = 'https://xeebucvvaxyglwjkkfvr.supabase.co/rest/v1';
// Server-side Supabase key. Never hardcode it here — this file is in git.
// Set the SUPABASE_SECRET_KEY script property (Project Settings > Script Properties) to an
// sb_secret_... key. The legacy SUPABASE_SERVICE_ROLE_KEY property is still read as a fallback.
const supabaseKeyAuth = (function() {
  const props = PropertiesService.getScriptProperties();
  return (props.getProperty('SUPABASE_SECRET_KEY') || props.getProperty('SUPABASE_SERVICE_ROLE_KEY') || '').trim();
})();

// New sb_secret_ keys are not JWTs and must only be sent on the apikey header;
// legacy JWT keys are also sent as the Bearer token.
function supabaseHeaders_(extra) {
  if (!supabaseKeyAuth) throw new Error("Supabase is not configured: set the SUPABASE_SECRET_KEY script property.");
  const headers = { 'apikey': supabaseKeyAuth };
  if (supabaseKeyAuth.indexOf('sb_') !== 0) headers['Authorization'] = 'Bearer ' + supabaseKeyAuth;
  return Object.assign(headers, extra || {});
}

// Checks a browser's Supabase access token with Supabase Auth and returns the staff user id.
// Throws unless the token belongs to a logged-in Office OS staff member.
function requireStaff_(accessToken) {
  const token = String(accessToken || '').trim();
  if (!token) throw new Error("Please log in again to use this feature.");
  const res = UrlFetchApp.fetch(supabaseUrlRest.replace('/rest/v1', '/auth/v1/user'), {
    'method': 'get',
    'headers': { 'apikey': supabaseKeyAuth, 'Authorization': 'Bearer ' + token },
    'muteHttpExceptions': true
  });
  if (res.getResponseCode() !== 200) throw new Error("Your session has expired. Please log in again.");
  const user = JSON.parse(res.getContentText());
  const staffId = user && user.app_metadata && user.app_metadata.app_user_id;
  if (!staffId) throw new Error("Only Office OS staff can use this feature.");
  return String(staffId);
}

// One-time migration scripts. The trailing underscore makes them private, so the public
// web app can no longer trigger them through google.script.run.
function migrateClientsToSupabase_() {
  const auditData = SpreadsheetApp.openById('1BVOTgOPsvGnpefO1zWeDeGwa-aLu3_qdm_kTE4JsUGY').getSheetByName('Clients').getDataRange().getValues();
  const debtData = SpreadsheetApp.openById('16jRkyPkZJygY6QsTU3L8urXQVZyGG-FvY9DHyuJv1-8').getSheetByName('Master').getDataRange().getValues();
  let uniqueClients = {};
  for (let i = 1; i < auditData.length; i++) { if (auditData[i][0]) uniqueClients[auditData[i][0].toString().trim()] = true; }
  for (let i = 1; i < debtData.length; i++) { if (debtData[i][0]) uniqueClients[debtData[i][0].toString().trim()] = true; }
  let payload = []; for (let clientName in uniqueClients) payload.push({ display_name: clientName, entity_type: "Migrated from Sheets" });
  if (payload.length > 0) UrlFetchApp.fetch(`${supabaseUrlRest}/clients`, { 'method' : 'post', 'contentType': 'application/json', 'headers': supabaseHeaders_({ 'Prefer': 'return=minimal' }), 'payload' : JSON.stringify(payload) });
}

// ==========================================
// 4. FIX SCRIPT: AUDIT USERS & HISTORY
// ==========================================
function fixAuditUsersAndHistory_() {
  const optionsGet = { 'method': 'get', 'headers': supabaseHeaders_() };
  const optionsPost = { 'method': 'post', 'contentType': 'application/json', 'headers': supabaseHeaders_({ 'Prefer': 'return=minimal' }) };

  // Helper function to normalize names for safer matching (safety net)
  function normalizeName(str) {
    if (!str) return "";
    return str.toString()
              .trim()
              .toLowerCase()
              .replace(/\(.*?\)/g, "")   // Strip parentheses
              .replace(/[^a-z0-9]/g, "") // Strip non-alphanumeric characters
              .replace(/\s+/g, "");      // Strip all spaces
  }

  // 1. Fetch Existing Task Logs to prevent duplicates without deleting anything (safety net)
  let existingLogs = [];
  try {
    existingLogs = JSON.parse(UrlFetchApp.fetch(supabaseUrlRest + '/task_logs?task_id=is.null&client_id=not.is.null&select=client_id,timestamp', optionsGet).getContentText());
  } catch (e) {
    Logger.log("Error fetching existing task logs: " + e.message);
  }
  
  const existingLogsSet = new Set();
  existingLogs.forEach(log => {
    if (log.client_id && log.timestamp) {
      const ms = new Date(log.timestamp).getTime();
      existingLogsSet.add(log.client_id + "_" + ms);
    }
  });
  
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

  // 6. Fetch Clients for mapping using safe name normalization
  const clients = JSON.parse(UrlFetchApp.fetch(supabaseUrlRest + '/clients?select=client_id,display_name', optionsGet).getContentText());
  const clientMap = {}; 
  clients.forEach(c => {
    const norm = normalizeName(c.display_name);
    if (norm) clientMap[norm] = c.client_id;
  });

  // 7. Fetch Work_Logs and migrate History
  const historySheet = SpreadsheetApp.openById(auditSheetId).getSheetByName('Work_Logs');
  const historyData = historySheet.getDataRange().getValues();
  
  // Column Match Safety Net
  if (!historyData || historyData.length < 2) {
    return "Error: Work_Logs sheet is empty or does not contain data.";
  }
  const headers = historyData[0];
  if (headers.length < 8) {
    return "Error: Work_Logs sheet must contain at least 8 columns. Found: " + headers.length;
  }

  let logsPayload = [];
  let unmatchedClients = new Set();
  let skippedCount = 0;
  
  // Track latest update values per client to update the clients table (safety net)
  const latestClientUpdate = {};
  
  for (let i = 1; i < historyData.length; i++) {
    let rawClientName = historyData[i][1] ? historyData[i][1].toString().trim() : null;
    if (!rawClientName) continue;
    let normClientName = normalizeName(rawClientName);
    let rawUserName = historyData[i][2] ? historyData[i][2].toString().trim().toLowerCase() : null;
    
    let clientId = clientMap[normClientName];
    
    let userId = userMap[rawUserName];
    if (!userId && rawUserName) {
       const foundUser = existingUsers.find(u => u.name.toLowerCase().includes(rawUserName));
       if(foundUser) userId = foundUser.id;
    }

    if (clientId) {
      let timestampVal;
      if (historyData[i][0]) {
        const d = new Date(historyData[i][0]);
        if (!isNaN(d.getTime())) {
          timestampVal = d.toISOString();
        } else {
          timestampVal = new Date().toISOString();
        }
      } else {
        timestampVal = new Date().toISOString();
      }
      
      const sheetMs = new Date(timestampVal).getTime();
      const compositeKey = clientId + "_" + sheetMs;
      
      // Update our "latest update" tracker (always do this so the clients table gets synced)
      if (!latestClientUpdate[clientId] || sheetMs > latestClientUpdate[clientId].time) {
        latestClientUpdate[clientId] = {
          time: sheetMs,
          bank: historyData[i][3] ? historyData[i][3].toString().trim() : '',
          sales: historyData[i][4] ? historyData[i][4].toString().trim() : '',
          ledger: historyData[i][5] ? historyData[i][5].toString().trim() : '',
          others: historyData[i][6] ? historyData[i][6].toString().trim() : '',
          note: historyData[i][7] ? historyData[i][7].toString().trim() : ''
        };
      }
      
      // Skip if this history log already exists in Supabase to prevent duplication
      if (existingLogsSet.has(compositeKey)) {
        skippedCount++;
        continue;
      }
      
      let rawNote = historyData[i][7] ? historyData[i][7].toString().trim() : '';
      let noteVal = rawNote ? '[Historical Audit] ' + rawNote : 'Historical Update';
      
      logsPayload.push({
        client_id: clientId,
        task_id: null,
        user_id: userId || null,
        timestamp: timestampVal,
        bank_status: historyData[i][3] ? historyData[i][3].toString().trim() : '',
        sales_status: historyData[i][4] ? historyData[i][4].toString().trim() : '',
        ledger_status: historyData[i][5] ? historyData[i][5].toString().trim() : '',
        others_status: historyData[i][6] ? historyData[i][6].toString().trim() : '',
        note: noteVal
      });
    } else {
      unmatchedClients.add(rawClientName);
    }
  }

  if (logsPayload.length > 0) {
    optionsPost.payload = JSON.stringify(logsPayload);
    UrlFetchApp.fetch(supabaseUrlRest + '/task_logs', optionsPost);
    Logger.log("Audit History migrated successfully. " + logsPayload.length + " new rows saved.");
  }
  
  // 8. Update client records in Supabase with their latest status from the sheet (safety net)
  let updatedClientsCount = 0;
  for (let clientId in latestClientUpdate) {
    const update = latestClientUpdate[clientId];
    
    // Determine audit status (Completed vs In progress)
    let statusVal = 'In progress';
    const lowerNote = update.note.toLowerCase();
    if (lowerNote.includes('completed') || lowerNote.includes('fully paid') || lowerNote.includes('done')) {
      statusVal = 'Completed';
    }
    
    const clientPayload = {
      audit_bank: update.bank,
      audit_sales: update.sales,
      audit_ledger: update.ledger,
      audit_others: update.others,
      audit_status: statusVal
    };
    
    const optionsPatch = {
      'method': 'patch',
      'contentType': 'application/json',
      'headers': supabaseHeaders_({ 'Prefer': 'return=minimal' }),
      'payload': JSON.stringify(clientPayload)
    };
    
    try {
      UrlFetchApp.fetch(supabaseUrlRest + '/clients?client_id=eq.' + clientId, optionsPatch);
      updatedClientsCount++;
    } catch (e) {
      Logger.log("Error updating client " + clientId + ": " + e.message);
    }
  }
  
  let unmatchedArray = Array.from(unmatchedClients);
  let statusMsg = "Audit History migration complete. Migrated: " + logsPayload.length + " new rows. Skipped: " + skippedCount + " duplicates. Synced board status for: " + updatedClientsCount + " active clients.";
  if (unmatchedArray.length > 0) {
    statusMsg += " Unmatched clients in sheet: " + unmatchedArray.join(", ");
  }
  return statusMsg;
}

// Sync push trigger (forced sync for .claspignore)
