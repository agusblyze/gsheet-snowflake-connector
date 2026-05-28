// ============================================================
// Snowflake → Google Sheets Connector
// Uses Snowflake SQL API v2 with Programmatic Access Token
// Supports saved queries, named tabs, and daily auto-refresh
// ============================================================

// ---------- CONFIGURATION ----------

function getConfig() {
  var props = PropertiesService.getUserProperties();
  return {
    account:   props.getProperty('SF_ACCOUNT')   || '',
    pat:       props.getProperty('SF_PAT')       || '',
    role:      props.getProperty('SF_ROLE')       || '',
    warehouse: props.getProperty('SF_WAREHOUSE')  || '',
    database:  props.getProperty('SF_DATABASE')   || '',
    schema:    props.getProperty('SF_SCHEMA')     || '',
  };
}

// ---------- MENU ----------

function buildSnowflakeMenu() {
  SpreadsheetApp.getUi()
    .createMenu('❄️ Snowflake')
    .addItem('🔍 Run Query', 'showSidebar')
    .addItem('🔄 Refresh All', 'refreshAllQueries')
    .addItem('⚙️ Configure', 'showConfig')
    .addToUi();
}

// ---------- SIDEBAR ----------

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('❄️ Snowflake Query')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showConfig() {
  var html = HtmlService.createHtmlOutputFromFile('config')
    .setTitle('⚙️ Snowflake Configuration')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------- SAVE / LOAD CONFIG (called from Config.html) ----------

function saveConfig(cfg) {
  var props = PropertiesService.getUserProperties();
  props.setProperties({
    SF_ACCOUNT:   cfg.account.trim(),
    SF_PAT:       cfg.pat.trim(),
    SF_WAREHOUSE: cfg.warehouse.trim(),
    SF_DATABASE:  cfg.database.trim(),
    SF_SCHEMA:    cfg.schema.trim(),
    SF_ROLE:      cfg.role.trim(),
  });
  return 'Configuration saved!';
}

function loadConfig() {
  var props = PropertiesService.getUserProperties();
  return {
    account:   props.getProperty('SF_ACCOUNT')   || 'A3617057978961-SYB22999',
    warehouse: props.getProperty('SF_WAREHOUSE')  || '',
    database:  props.getProperty('SF_DATABASE')   || '',
    schema:    props.getProperty('SF_SCHEMA')     || '',
    role:      props.getProperty('SF_ROLE')       || 'ANALYTICS_ROLE',
  };
}

// ============================================================
// SAVED QUERIES — stored as JSON in document properties
// so they're tied to THIS spreadsheet and shared with editors
// ============================================================

function getSavedQueries() {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty('SF_SAVED_QUERIES');
  if (!raw) return [];
  try {
    return JSON.parse(raw);
  } catch (e) {
    return [];
  }
}

function saveQueryList(queries) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('SF_SAVED_QUERIES', JSON.stringify(queries));
}

// Called from sidebar: save or update a query
function saveQuery(name, sql, autoRefresh) {
  if (!name || !name.trim()) throw new Error('Query name is required.');
  if (!sql || !sql.trim()) throw new Error('SQL is required.');

  var queries = getSavedQueries();
  var found = false;
  for (var i = 0; i < queries.length; i++) {
    if (queries[i].name === name.trim()) {
      queries[i].sql = sql;
      queries[i].autoRefresh = !!autoRefresh;
      found = true;
      break;
    }
  }
  if (!found) {
    queries.push({
      name: name.trim(),
      sql: sql,
      autoRefresh: !!autoRefresh,
    });
  }
  saveQueryList(queries);
  syncDailyTrigger();
  return queries;
}

// Called from sidebar: delete a saved query
function deleteQuery(name) {
  var queries = getSavedQueries();
  var filtered = [];
  for (var i = 0; i < queries.length; i++) {
    if (queries[i].name !== name) filtered.push(queries[i]);
  }
  saveQueryList(filtered);
  syncDailyTrigger();
  return filtered;
}

// Called from sidebar: toggle auto-refresh for a query
function toggleAutoRefresh(name, enabled) {
  var queries = getSavedQueries();
  for (var i = 0; i < queries.length; i++) {
    if (queries[i].name === name) {
      queries[i].autoRefresh = !!enabled;
      break;
    }
  }
  saveQueryList(queries);
  syncDailyTrigger();
  return queries;
}

// ============================================================
// DAILY TRIGGER MANAGEMENT
// ============================================================

function syncDailyTrigger() {
  // Check if any query has autoRefresh enabled
  var queries = getSavedQueries();
  var needsTrigger = false;
  for (var i = 0; i < queries.length; i++) {
    if (queries[i].autoRefresh) {
      needsTrigger = true;
      break;
    }
  }

  // Remove existing daily triggers for refreshAllQueries
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshAllQueries') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a new daily trigger if needed (runs at ~6 AM)
  if (needsTrigger) {
    ScriptApp.newTrigger('refreshAllQueries')
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();
  }
}

// Refresh all queries that have autoRefresh enabled
function refreshAllQueries() {
  var queries = getSavedQueries();
  var errors = [];

  for (var i = 0; i < queries.length; i++) {
    if (queries[i].autoRefresh) {
      try {
        var result = runSnowflakeQuery(queries[i].sql);
        writeResultsToSheet(result, queries[i].name);
      } catch (e) {
        errors.push(queries[i].name + ': ' + e.message);
      }
    }
  }

  if (errors.length > 0) {
    Logger.log('Refresh errors: ' + errors.join('; '));
  }
}

// ============================================================
// QUERY EXECUTION
// ============================================================

function runSnowflakeQuery(sql) {
  var cfg = getConfig();
  if (!cfg.account || !cfg.pat) {
    throw new Error('Snowflake credentials not configured. Use ❄️ Snowflake > ⚙️ Configure to set your Account and PAT.');
  }

  var baseUrl = 'https://' + cfg.account + '.snowflakecomputing.com';
  var submitUrl = baseUrl + '/api/v2/statements';

  var body = {
    statement: sql,
    timeout: 60,
  };
  if (cfg.role)      body.role      = cfg.role;
  if (cfg.warehouse) body.warehouse = cfg.warehouse;
  if (cfg.database)  body.database  = cfg.database;
  if (cfg.schema)    body.schema    = cfg.schema;

  var headers = {
    'Authorization': 'Bearer ' + cfg.pat,
    'X-Snowflake-Authorization-Token-Type': 'PROGRAMMATIC_ACCESS_TOKEN',
    'Accept': 'application/json',
  };

  var submitResp = UrlFetchApp.fetch(submitUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });

  var httpCode = submitResp.getResponseCode();
  var respText = submitResp.getContentText();

  if (!respText || respText.trim() === '') {
    throw new Error('Snowflake returned empty response (HTTP ' + httpCode + ').');
  }

  var data;
  try {
    data = JSON.parse(respText);
  } catch (e) {
    throw new Error('Snowflake non-JSON response (HTTP ' + httpCode + '): ' + respText.substring(0, 300));
  }

  // Handle async: poll if 202
  if (httpCode === 202) {
    var statusUrl = data.statementStatusUrl
      ? baseUrl + data.statementStatusUrl
      : submitUrl + '/' + data.statementHandle;

    for (var poll = 0; poll < 60; poll++) {
      Utilities.sleep(1000);
      var pollResp = UrlFetchApp.fetch(statusUrl, {
        method: 'get',
        headers: headers,
        muteHttpExceptions: true,
      });
      httpCode = pollResp.getResponseCode();
      data = JSON.parse(pollResp.getContentText());
      if (httpCode === 200) break;
    }
  }

  if (httpCode !== 200) {
    var errMsg = data.message || data.error || JSON.stringify(data);
    throw new Error('Snowflake error (HTTP ' + httpCode + '): ' + errMsg);
  }

  // Extract columns and types
  var columns = [];
  var colTypes = [];
  if (data.resultSetMetaData && data.resultSetMetaData.rowType) {
    for (var i = 0; i < data.resultSetMetaData.rowType.length; i++) {
      columns.push(data.resultSetMetaData.rowType[i].name);
      colTypes.push(data.resultSetMetaData.rowType[i].type);
    }
  }

  // Collect rows from partition 0 (already in the response)
  var rawRows = data.data || [];

  // Fetch remaining partitions if any
  var partitionInfo = data.resultSetMetaData && data.resultSetMetaData.partitionInfo
    ? data.resultSetMetaData.partitionInfo
    : [];
  var statementHandle = data.statementHandle;

  if (partitionInfo.length > 1 && statementHandle) {
    for (var p = 1; p < partitionInfo.length; p++) {
      var partUrl = submitUrl + '/' + statementHandle + '?partition=' + p;
      var partResp = UrlFetchApp.fetch(partUrl, {
        method: 'get',
        headers: headers,
        muteHttpExceptions: true,
      });
      var partCode = partResp.getResponseCode();
      if (partCode === 200) {
        var partData = JSON.parse(partResp.getContentText());
        var partRows = partData.data || [];
        for (var pr = 0; pr < partRows.length; pr++) {
          rawRows.push(partRows[pr]);
        }
      }
    }
  }

  // Convert values based on column types
  var rows = [];
  var EPOCH = new Date(1970, 0, 1);
  for (var r = 0; r < rawRows.length; r++) {
    var row = [];
    for (var c = 0; c < rawRows[r].length; c++) {
      var val = rawRows[r][c];
      var typ = colTypes[c] || '';

      if (val === null || val === undefined) {
        row.push('');
      } else if (typ === 'date') {
        var d = new Date(EPOCH.getTime() + Number(val) * 86400000);
        row.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
      } else if (typ === 'timestamp_ntz' || typ === 'timestamp_ltz' || typ === 'timestamp_tz') {
        var ts = Number(val);
        if (ts > 1e15) ts = ts / 1e6;
        else if (ts < 1e12) ts = ts * 1000;
        var dt = new Date(ts);
        row.push(Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      } else if (typ === 'fixed' || typ === 'real' || typ === 'float') {
        row.push(Number(val));
      } else {
        row.push(val);
      }
    }
    rows.push(row);
  }

  return { columns: columns, rows: rows };
}

// ============================================================
// WRITE RESULTS TO A NAMED SHEET TAB
// ============================================================

function writeResultsToSheet(result, tabName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Truncate tab name to 100 chars (Sheets limit)
  tabName = tabName.substring(0, 100);

  // Reuse existing tab or create new one
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  // Write header row
  if (result.columns && result.columns.length > 0) {
    sheet.getRange(1, 1, 1, result.columns.length).setValues([result.columns]);
    var headerRange = sheet.getRange(1, 1, 1, result.columns.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4A90D9');
    headerRange.setFontColor('#FFFFFF');
  }

  // Write data rows in chunks
  var chunkSize = 500;
  for (var i = 0; i < result.rows.length; i += chunkSize) {
    var chunk = result.rows.slice(i, i + chunkSize);
    if (chunk.length > 0) {
      sheet.getRange(
        sheet.getLastRow() + 1,
        1,
        chunk.length,
        chunk[0].length
      ).setValues(chunk);
    }
  }

  if (result.columns.length > 0) {
    sheet.autoResizeColumns(1, result.columns.length);
  }

  // Add a small "last refreshed" note in the tab
  var lastRefreshed = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var noteCell = sheet.getRange(1, result.columns.length + 2);
  noteCell.setValue('Last refreshed: ' + lastRefreshed);
  noteCell.setFontColor('#9aa0a6');
  noteCell.setFontSize(9);

  ss.setActiveSheet(sheet);

  return {
    tabName: tabName,
    rowCount: result.rows.length,
    colCount: result.columns.length,
  };
}

// ============================================================
// ENTRY POINTS (called from sidebar)
// ============================================================

// Run a query and write to a named tab
function executeQuery(name, sql) {
  if (!name || name.trim() === '') {
    throw new Error('Please enter a query name.');
  }
  if (!sql || sql.trim() === '') {
    throw new Error('Please enter a SQL query.');
  }
  var result = runSnowflakeQuery(sql);
  var info = writeResultsToSheet(result, name.trim());
  return '✅ Done! ' + info.rowCount + ' rows × ' + info.colCount + ' columns written to tab "' + info.tabName + '".';
}

// Run a saved query by name (used by refresh and sidebar)
function runSavedQuery(name) {
  var queries = getSavedQueries();
  for (var i = 0; i < queries.length; i++) {
    if (queries[i].name === name) {
      var result = runSnowflakeQuery(queries[i].sql);
      var info = writeResultsToSheet(result, queries[i].name);
      return '✅ Done! ' + info.rowCount + ' rows × ' + info.colCount + ' columns written to tab "' + info.tabName + '".';
    }
  }
  throw new Error('Saved query "' + name + '" not found.');
}

// ---------- TRIGGER INSTALLER (run once) ----------
function installSnowflakeTrigger() {
  var ssId = '115bLtRj45Xq68N6gJVB_6wPjg3GYwj7Zk9boBeH3fK0';
  var existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(t){
    if (t.getHandlerFunction() === 'buildSnowflakeMenu') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildSnowflakeMenu')
    .forSpreadsheet(SpreadsheetApp.openById(ssId))
    .onOpen()
    .create();
  Logger.log('Snowflake onOpen trigger installed for spreadsheet ' + ssId);
}
