// ── Influencer Tracker — Apps Script Server ──────────────────────────────────
// Paste this entire file into your Apps Script editor as Code.gs
// Then paste the contents of influencer_tracker.html as a new HTML file named "index"
// Deploy → New deployment → Web app → Execute as: Me → Access: Anyone in your org

function doGet(e) {
  // JSON API endpoint for local-mode sync (?action=getData)
  if (e && e.parameter && e.parameter.action === 'getData') {
    var callback = e.parameter.callback || '';
    var logResult = getSheetRows('Outreach_Log');
    if (logResult.error) logResult = { rows: [] };
    var result = {
      paid:         getSheetRows('Paid_Collabs'),
      ugc:          getSheetRows('UGC_Free'),
      posts:        getSheetRows('Posts'),
      outreach:     getSheetRows('Outreach'),
      outreach_log: logResult
    };
    var json = JSON.stringify(result);
    return ContentService
      .createTextOutput(callback ? callback + '(' + json + ')' : json)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Influencer Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPaidData() {
  return getSheetRows('Paid_Collabs');
}

function getUGCData() {
  return getSheetRows('UGC_Free');
}

function getPostsData() {
  return getSheetRows('Posts');
}

function getTemplatesData() {
  return getSheetRows('Outreach');
}

function getOutreachLogData() {
  var result = getSheetRows('Outreach_Log');
  if (result.error) return { rows: [] };
  return result;
}

function saveTemplate(tpl) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Outreach');
    if (!sheet) return { error: 'Tab "Outreach" not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var idCol   = headers.indexOf('id');

    function buildRow(id) {
      return headers.map(function(h) {
        if (h === 'id')      return id;
        if (h === 'type')    return tpl.type    || '';
        if (h === 'title')   return tpl.title   || '';
        if (h === 'content') return tpl.content || '';
        return '';
      });
    }

    if (tpl.id) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) === String(tpl.id)) {
          sheet.getRange(i + 1, 1, 1, headers.length).setValues([buildRow(tpl.id)]);
          return { success: true, id: tpl.id };
        }
      }
    }

    var maxId = 0;
    for (var i = 1; i < data.length; i++) {
      var n = parseInt(data[i][idCol]);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
    var newId = maxId + 1;
    sheet.appendRow(buildRow(newId));
    return { success: true, id: newId };
  } catch(e) {
    return { error: e.message };
  }
}

function deleteTemplate(id) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Outreach');
    if (!sheet) return { error: 'Tab "Outreach" not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var idCol   = headers.indexOf('id');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { error: 'Template id ' + id + ' not found.' };
  } catch(e) {
    return { error: e.message };
  }
}

function updatePaidRow(handle, fields) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Paid_Collabs');
    if (!sheet) return { error: 'Tab "Paid_Collabs" not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) {
      return String(h).trim().toLowerCase().replace(/[\s\/]+/g, '_');
    });
    var handleCol = headers.indexOf('handle');
    if (handleCol < 0) return { error: 'Column "handle" not found.' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][handleCol]).trim().toLowerCase() === String(handle).trim().toLowerCase()) {
        var keys = Object.keys(fields);
        for (var k = 0; k < keys.length; k++) {
          var col = headers.indexOf(keys[k]);
          if (col >= 0) {
            sheet.getRange(i + 1, col + 1).setValue(fields[keys[k]] || '');
          }
        }
        return { success: true };
      }
    }
    return { error: 'Handle "' + handle + '" not found in sheet.' };
  } catch(e) {
    return { error: e.message };
  }
}

function updateUGCRow(handle, fields) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('UGC_Free');
    if (!sheet) return { error: 'Tab "UGC_Free" not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) {
      return String(h).trim().toLowerCase().replace(/[\s\/]+/g, '_');
    });
    var handleCol = headers.indexOf('handle');
    if (handleCol < 0) return { error: 'Column "handle" not found.' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][handleCol]).trim().toLowerCase() === String(handle).trim().toLowerCase()) {
        var keys = Object.keys(fields);
        for (var k = 0; k < keys.length; k++) {
          var col = headers.indexOf(keys[k]);
          if (col >= 0) {
            sheet.getRange(i + 1, col + 1).setValue(fields[keys[k]] || '');
          }
        }
        return { success: true };
      }
    }
    return { error: 'Handle "' + handle + '" not found in sheet.' };
  } catch(e) {
    return { error: e.message };
  }
}

function addPaidRow(handle, month, type, email) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Paid_Collabs');
    if (!sheet) return { error: 'Tab "Paid_Collabs" not found.' };
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) {
      return String(h).trim().toLowerCase().replace(/[\s\/]+/g, '_');
    });
    var row = headers.map(function(h) {
      if (h === 'handle')   return handle || '';
      if (h === 'month')    return month  || '';
      if (h === 'type')     return type   || '';
      if (h === 'email' || h === 'email_dm') return email || '';
      if (h === 'status')   return 'Initial Contact';
      return '';
    });
    sheet.appendRow(row);
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

function addUGCRow(handle, month, shipDate, product) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('UGC_Free');
    if (!sheet) return { error: 'Tab "UGC_Free" not found.' };
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) {
      return String(h).trim().toLowerCase().replace(/[\s\/]+/g, '_');
    });
    var row = headers.map(function(h) {
      if (h === 'handle')    return handle   || '';
      if (h === 'month')     return month    || '';
      if (h === 'ship_date') return shipDate || '';
      if (h === 'product')   return product  || '';
      if (h === 'status')    return 'Initial Contact';
      return '';
    });
    sheet.appendRow(row);
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

function saveAllPosts(posts) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Posts');
    if (!sheet) return { error: 'Tab "Posts" not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) {
      return String(h).trim().toLowerCase().replace(/[\s\/]+/g, '_');
    });

    // Clear existing data rows (preserve header)
    if (data.length > 1) {
      sheet.getRange(2, 1, data.length - 1, headers.length).clearContent();
    }
    if (!posts || !posts.length) return { success: true };

    // Map post fields to sheet column order
    var fieldAliases = {
      'date': 'scheduled_date',
      'scheduled_date': 'scheduled_date',
      'collab_handle': 'collab_handle',
      'collab handle': 'collab_handle',
      'caption_ready': 'caption_ready',
      'caption ready': 'caption_ready',
      'post_link': 'post_link',
      'post link': 'post_link'
    };
    var rows = posts.map(function(p) {
      return headers.map(function(h) {
        var field = fieldAliases[h] || h;
        return p[field] !== undefined ? (p[field] || '') : (p[h] || '');
      });
    });

    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    return { success: true, count: rows.length };
  } catch(e) {
    return { error: e.message };
  }
}

function saveAllOutreachLog(entries) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Outreach_Log');

    // Create sheet with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Outreach_Log');
      sheet.getRange(1, 1, 1, 6).setValues([['date', 'handle', 'template', 'channel', 'notes', 'follow_up_date']]);
    }

    var headers = ['date', 'handle', 'template', 'channel', 'notes', 'follow_up_date'];
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
    }
    if (!entries || !entries.length) return { success: true };

    var rows = entries.map(function(e) {
      return headers.map(function(h) { return e[h] || ''; });
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    return { success: true, count: rows.length };
  } catch(e) {
    return { error: e.message };
  }
}

function getSheetRows(sheetName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: 'Tab "' + sheetName + '" not found. Check the name is exactly correct (case-sensitive).' };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { rows: [] };

    var headers = data[0].map(function(h) {
      return String(h).trim().toLowerCase().replace(/[\s\/]+/g, '_');
    });

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // skip blank rows
      var obj = {};
      headers.forEach(function(h, j) {
        var v = row[j];
        if (v instanceof Date) {
          obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          obj[h] = String(v === null || v === undefined ? '' : v);
        }
      });
      rows.push(obj);
    }
    return { rows: rows };
  } catch (e) {
    return { error: e.message };
  }
}
