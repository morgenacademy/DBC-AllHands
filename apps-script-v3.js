function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);

  // ── E3 AGENDAPUNT ──
  if (data.type === 'e3') {
    var eS = ss.getSheetByName('E3');
    if (!eS) {
      eS = ss.insertSheet('E3');
      eS.appendRow(['Timestamp','Meeting','Vraag','Doel','Eigenaar','Gewenst resultaat','Link','Tijd','Conclusie']);
    }
    eS.appendRow([
      new Date().toISOString(),
      data.meeting || '',
      data.vraag || '',
      data.doel || '',
      data.eigenaar || '',
      data.resultaat || '',
      data.link || '',
      data.tijd || '5',
      data.conclusie || ''
    ]);
    return ContentService.createTextOutput('ok');
  }

  // ── E3 CONCLUSIES OPSLAAN (bulk update vanuit "overleg afgelopen") ──
  if (data.type === 'e3_conclusies') {
    var cS = ss.getSheetByName('E3');
    if (!cS) return ContentService.createTextOutput('no sheet');
    var allData = cS.getDataRange().getValues();
    (data.conclusies || []).forEach(function(item) {
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][1]).trim() === String(item.meeting).trim() &&
            String(allData[i][2]).trim() === String(item.vraag).trim()) {
          cS.getRange(i + 1, 9).setValue(item.conclusie || '');
          break;
        }
      }
    });
    return ContentService.createTextOutput('ok');
  }

  // ── NIEUW ACTIEPUNT (vanuit commitment of E3) ──
  if (data.type === 'taak') {
    var tS = ss.getSheetByName('Actiepunten');
    tS.appendRow([
      new Date().toISOString(),
      data.wie || '',
      data.taak || '',
      data.dod || '',
      data.deadline || '',
      '',
      data.maand || '',
      data.hulp_nodig || '',
      data.hulp_toel || ''
    ]);
    return ContentService.createTextOutput('ok');
  }

  // ── ACTIEPUNT STATUS UPDATES (done/not-done vanuit accountability) ──
  if (data.type === 'actiepunt_update') {
    var aS = ss.getSheetByName('Actiepunten');
    var allData = aS.getDataRange().getValues();
    var log = [];
    (data.items || []).forEach(function(item) {
      var found = false;
      var wie = String(item.wie || '').trim();
      var taak = String(item.taak || '').trim();
      for (var i = 1; i < allData.length; i++) {
        var sheetWie = String(allData[i][1]).trim();
        var sheetTaak = String(allData[i][2]).trim();
        if (sheetWie === wie && sheetTaak === taak) {
          aS.getRange(i + 1, 6).setValue(item.status);
          if (item.toelichting) aS.getRange(i + 1, 9).setValue(item.toelichting);
          if (item.deadline) aS.getRange(i + 1, 5).setValue(item.deadline);
          found = true;
          log.push('MATCH rij ' + (i + 1) + ': ' + sheetWie + ' / ' + sheetTaak.substring(0, 40));
          break;
        }
      }
      if (!found) {
        log.push('GEEN MATCH: "' + wie + '" / "' + taak.substring(0, 40) + '"');
      }
    });
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('Log');
      logSheet.appendRow(['Timestamp', 'Type', 'Persoon', 'Details']);
    }
    logSheet.appendRow([new Date().toISOString(), 'actiepunt_update', data.persoon || '', log.join(' | ')]);
    return ContentService.createTextOutput('ok');
  }

  // ── DELETE UPDATE (verwijder project update uit Sheet) ──
  if (data.type === 'delete_update') {
    var dS = ss.getSheetByName('Updates');
    var allData = dS.getDataRange().getValues();
    var persoon = String(data.persoon || '').trim();
    var stroom = String(data.stroom || '').trim();
    var project = String(data.project || '').trim();
    var maand = String(data.maand || '').trim();
    // Verwijder ALLE rijen die matchen (van onder naar boven om rij-indexen correct te houden)
    var rowsToDelete = [];
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][2]).trim() === persoon &&
          String(allData[i][3]).trim() === stroom &&
          String(allData[i][4]).trim() === project &&
          String(allData[i][1]).trim() === maand) {
        rowsToDelete.push(i + 1); // +1 voor sheet rij (1-based)
      }
    }
    // Verwijder van onder naar boven
    for (var j = rowsToDelete.length - 1; j >= 0; j--) {
      dS.deleteRow(rowsToDelete[j]);
    }
    // Log
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('Log');
      logSheet.appendRow(['Timestamp', 'Type', 'Persoon', 'Details']);
    }
    logSheet.appendRow([new Date().toISOString(), 'delete_update', persoon, 'Deleted ' + rowsToDelete.length + ' rows: ' + stroom + ' / ' + project + ' / ' + maand]);
    return ContentService.createTextOutput('ok');
  }

  // ── REGULIERE PROJECT UPDATES ──
  var uS = ss.getSheetByName('Updates');
  if (data.updates && data.updates.length) {
    data.updates.forEach(function(u) {
      uS.appendRow([
        new Date().toISOString(),
        data.maand || '',
        data.persoon || '',
        u.stroom || '',
        u.project || '',
        u.update || '',
        u.status || '',
        u.target_ytd || '',
        u.doel || '',
        u.partner || ''
      ]);
    });
  }
  return ContentService.createTextOutput('ok');
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uS = ss.getSheetByName('Updates');
  var data = uS.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var stroom = data[i][3];
    var project = data[i][4];
    if (stroom && project) {
      if (!result[stroom]) result[stroom] = [];
      if (result[stroom].indexOf(project) === -1) result[stroom].push(project);
    }
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════
// CLEANUP FUNCTIE — handmatig uitvoeren vanuit Script Editor
// Menu: Run > cleanupDuplicates
// ═══════════════════════════════════════════════════════════
function cleanupDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var totalRemoved = 0;

  // ── 1. UPDATES: dedup op Persoon(2) + Stroom(3) + Project(4) + Maand(1) ──
  var uS = ss.getSheetByName('Updates');
  if (uS) {
    var uData = uS.getDataRange().getValues();
    var uSeen = {};
    var uKeep = [0]; // altijd header behouden (rij-index 0)
    // Loop vooruit, onthoud LAATSTE index per key
    for (var i = 1; i < uData.length; i++) {
      var key = String(uData[i][2]).trim() + '|' + String(uData[i][3]).trim() + '|' + String(uData[i][4]).trim() + '|' + String(uData[i][1]).trim();
      uSeen[key] = i;
    }
    // Bouw set van te behouden rij-indices
    var uKeepSet = {};
    uKeepSet[0] = true; // header
    for (var k in uSeen) { uKeepSet[uSeen[k]] = true; }
    // Verwijder van onder naar boven
    var uRemoved = 0;
    for (var i = uData.length - 1; i >= 1; i--) {
      if (!uKeepSet[i]) {
        uS.deleteRow(i + 1);
        uRemoved++;
      }
    }
    totalRemoved += uRemoved;
    Logger.log('Updates: ' + uRemoved + ' duplicaten verwijderd, ' + Object.keys(uKeepSet).length + ' rijen behouden');
  }

  // ── 2. E3: dedup op Meeting(1) + Vraag(2) ──
  var eS = ss.getSheetByName('E3');
  if (eS) {
    var eData = eS.getDataRange().getValues();
    var eSeen = {};
    for (var i = 1; i < eData.length; i++) {
      var key = String(eData[i][1]).trim() + '|' + String(eData[i][2]).trim();
      eSeen[key] = i;
    }
    var eKeepSet = {};
    eKeepSet[0] = true;
    for (var k in eSeen) { eKeepSet[eSeen[k]] = true; }
    var eRemoved = 0;
    for (var i = eData.length - 1; i >= 1; i--) {
      if (!eKeepSet[i]) {
        eS.deleteRow(i + 1);
        eRemoved++;
      }
    }
    totalRemoved += eRemoved;
    Logger.log('E3: ' + eRemoved + ' duplicaten verwijderd');
  }

  // ── 3. ACTIEPUNTEN: dedup op Wie(1) + Taak(2) ──
  var aS = ss.getSheetByName('Actiepunten');
  if (aS) {
    var aData = aS.getDataRange().getValues();
    var aSeen = {};
    for (var i = 1; i < aData.length; i++) {
      var key = String(aData[i][1]).trim() + '|' + String(aData[i][2]).trim();
      aSeen[key] = i;
    }
    var aKeepSet = {};
    aKeepSet[0] = true;
    for (var k in aSeen) { aKeepSet[aSeen[k]] = true; }
    var aRemoved = 0;
    for (var i = aData.length - 1; i >= 1; i--) {
      if (!aKeepSet[i]) {
        aS.deleteRow(i + 1);
        aRemoved++;
      }
    }
    totalRemoved += aRemoved;
    Logger.log('Actiepunten: ' + aRemoved + ' duplicaten verwijderd');
  }

  // Resultaat
  var msg = 'Cleanup voltooid: ' + totalRemoved + ' dubbele rijen verwijderd.';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}
