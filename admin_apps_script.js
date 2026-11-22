var SHEET_ID = '1Qv_TxvPY0_64ZVbVG2s8ZLvLyFJxtctCKsjealm23fE';
var DEFAULT_LEGACY_TOURNAMENT_ID = 'VB-2026';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sports Portal Admin Panel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function matchesTournament(cellValue, targetId) {
  var val = String(cellValue || "").trim();
  var target = String(targetId || "").trim();
  if (val === target) return true;
  if (val === "" && target === DEFAULT_LEGACY_TOURNAMENT_ID) return true;
  return false;
}

function getTournaments() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  if (!sheet) {
    sheet = ss.insertSheet("Tournaments");
    sheet.appendRow([
      "Tournament ID", "Sport", "Year", "Tournament Name", "Status",
      "Deadline", "Logo URL", "Rules", "Max Players", "Max Officials"
    ]);
    sheet.getRange("A1:J1").setFontWeight("bold");
    sheet.appendRow([DEFAULT_LEGACY_TOURNAMENT_ID, "Volleyball", 2026, "RCI Volleyball Tournament 2026", "Active", "", "", "", 14, 3]);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var list = [];
  
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var formattedDeadline = "";
    if (data[i][5] instanceof Date) {
      var tzoffset = (new Date()).getTimezoneOffset() * 60000;
      formattedDeadline = (new Date(data[i][5] - tzoffset)).toISOString().slice(0, 16);
    } else if (data[i][5]) {
      formattedDeadline = String(data[i][5]);
    }
    
    list.push({
      row: i + 2,
      id: String(data[i][0]).trim(),
      sport: String(data[i][1] || "Volleyball").trim(),
      year: String(data[i][2] || "").trim(),
      name: String(data[i][3] || "").trim(),
      status: String(data[i][4] || "Active").trim(),
      deadline: formattedDeadline,
      logo: String(data[i][6] || "").trim(),
      rules: String(data[i][7] || ""),
      maxPlayers: data[i][8] || 14,
      maxOfficials: data[i][9] || 3
    });
  }
  
  list.sort(function(a, b) {
    return (parseInt(b.year) || 0) - (parseInt(a.year) || 0);
  });
  return list;
}

function createTournament(data) {
  if (!data) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  if (!sheet) {
    getTournaments();
    sheet = ss.getSheetByName("Tournaments");
  }
  
  var prefix = data.sport === "Futsal" ? "FS" : (data.sport === "Handball" ? "HB" : "VB");
  var newId = prefix + "-" + data.year + "-" + Math.floor(100 + Math.random() * 900);
  
  var logoUrl = data.logoLink ? data.logoLink.trim() : "";
  if (!logoUrl && data.logoBase64) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(data.logoBase64), data.logoMimeType, data.logoName);
      var file = DriveApp.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
      logoUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    } catch (e) {
      return "Error uploading logo: " + e.message;
    }
  }
  
  sheet.appendRow([
    newId,
    data.sport,
    data.year,
    data.name.trim(),
    data.status || "Active",
    data.deadline || "",
    logoUrl,
    data.rules || "",
    data.maxPlayers || 14,
    data.maxOfficials || 3
  ]);
  
  return "New tournament '" + data.name + "' (" + data.year + ") created successfully!";
}

function updateTournamentStatus(row, status) {
  if (!row || !status) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  sheet.getRange(row, 5).setValue(status);
  return "Tournament status updated to: " + status;
}

function updateTournamentSettings(data) {
  if (!data) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  if (!sheet) return "Error: Tournaments sheet not found.";
  
  var all = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]).trim() === String(data.tournamentId).trim()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return "Error: Tournament not found.";
  
  sheet.getRange(rowIndex, 4).setValue(data.name);
  sheet.getRange(rowIndex, 5).setValue(data.status);
  sheet.getRange(rowIndex, 6).setValue(data.deadline);
  sheet.getRange(rowIndex, 8).setValue(data.rules);
  sheet.getRange(rowIndex, 9).setValue(data.maxPlayers);
  sheet.getRange(rowIndex, 10).setValue(data.maxOfficials);

  if (data.logoLink && data.logoLink.trim() !== "") {
    sheet.getRange(rowIndex, 7).setValue(data.logoLink.trim());
  } else if (data.logoBase64) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(data.logoBase64), data.logoMimeType, data.logoName);
      var file = DriveApp.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
      sheet.getRange(rowIndex, 7).setValue("https://lh3.googleusercontent.com/d/" + file.getId());
    } catch (e) {
      return "Settings saved, but logo upload failed: " + e.message;
    }
  }
  return "Tournament settings updated!";
}

function getTeamsForApproval(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var numCols = Math.max(sheet.getLastColumn(), 9);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var teams = [];
  
  for (var i = 0; i < data.length; i++) {
    if (!data[i][1]) continue;
    if (!matchesTournament(data[i][8], tournamentId)) continue;
    teams.push({
      row: i + 2,
      name: String(data[i][1]),
      captain: String(data[i][2]),
      division: String(data[i][6] || "Men's"),
      status: String(data[i][7] || "Pending")
    });
  }
  return teams;
}

function updateTeamStatus(row, status) {
  if (!row || !status) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  sheet.getRange(row, 8).setValue(status);
  return "Team status updated to: " + status;
}

function getTeams(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var numCols = Math.max(sheet.getLastColumn(), 9);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var list = [];
  
  for (var i = 0; i < data.length; i++) {
    if (!data[i][1]) continue;
    if (!matchesTournament(data[i][8], tournamentId)) continue;
    list.push(String(data[i][1]).trim());
  }
  return list;
}

function updatePointsTable(data) {
  if (!data) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Points Table");
  if (!sheet) {
    sheet = ss.insertSheet("Points Table");
    sheet.appendRow(["Rank", "Team", "PLD", "W", "L", "PTS", "SW/GF", "SL/GA", "SDiff/GD", "PW/D", "PL", "PDiff", "Logo", "Division", "Tournament ID"]);
  }
  
  var regSheet = ss.getSheetByName("Registration");
  var logoUrl = "No Logo";
  var division = "Men's";
  if (regSheet) {
    var regData = regSheet.getDataRange().getValues();
    for (var j = 1; j < regData.length; j++) {
      if (matchesTournament(regData[j][8], data.tournamentId) && String(regData[j][1]).trim() === String(data.teamName).trim()) {
        logoUrl = regData[j][5];
        division = regData[j][6];
        break;
      }
    }
  }
  
  var rowData = [
    "-", data.teamName, data.mPld, data.mW, data.mL, data.mPts,
    data.sW, data.sL, data.sDiff,
    data.pW, data.pL, data.pDiff,
    logoUrl, division, data.tournamentId
  ];
  
  var lastRow = sheet.getLastRow();
  var rowIndex = -1;
  if (lastRow >= 2) {
    var numCols = Math.max(sheet.getLastColumn(), 15);
    var dataRange = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    for (var i = 0; i < dataRange.length; i++) {
      if (matchesTournament(dataRange[i][14], data.tournamentId) && String(dataRange[i][1]).trim() === String(data.teamName).trim()) {
        rowIndex = i + 2;
        break;
      }
    }
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 15).setValues([rowData]);
    return "Standings updated for " + data.teamName;
  } else {
    sheet.appendRow(rowData);
    return "New standing added for " + data.teamName;
  }
}

function addFixture(data) {
  if (!data) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Fixtures");
  if (!sheet) {
    sheet = ss.insertSheet("Fixtures");
    sheet.appendRow(["Match ID", "Date", "Time", "Team A", "Team B", "Venue", "Status", "Result", "Tournament ID"]);
  }
  sheet.appendRow([
    data.matchId, data.date, data.time,
    data.teamA, data.teamB, data.venue,
    data.status, data.result, data.tournamentId
  ]);
  if (String(data.status || "").toLowerCase() === "completed" && data.result) {
    recalculatePointsTable(data.tournamentId);
    return "Match saved & Points Table automatically updated!";
  }
  return "Match scheduled successfully!";
}

function getFixturesForAdmin(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Fixtures");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numCols = Math.max(sheet.getLastColumn(), 9);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var tz = Session.getScriptTimeZone();
  var list = [];

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0] && !data[i][3]) continue;
    if (!matchesTournament(data[i][8], tournamentId)) continue;

    var dStr = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], tz, "yyyy-MM-dd") : String(data[i][1] || "");
    var tStr = data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], tz, "HH:mm") : String(data[i][2] || "");

    list.push({
      row: i + 2,
      matchId: String(data[i][0] || ""),
      date: dStr,
      time: tStr,
      teamA: String(data[i][3] || ""),
      teamB: String(data[i][4] || ""),
      venue: String(data[i][5] || ""),
      status: String(data[i][6] || "Upcoming"),
      result: String(data[i][7] || "")
    });
  }
  return list;
}

function updateMatchScore(data) {
  if (!data || !data.row) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Fixtures");
  if (!sheet) return "Error: Fixtures sheet not found.";

  var resultStr = "-";
  var sA = String(data.scoreA !== undefined && data.scoreA !== null ? data.scoreA : "").trim();
  var sB = String(data.scoreB !== undefined && data.scoreB !== null ? data.scoreB : "").trim();
  var details = String(data.setDetails || "").trim();

  if (sA !== "" && sB !== "") {
    resultStr = sA + " - " + sB;
    if (details !== "") {
      resultStr += " (" + details + ")";
    }
  } else if (data.result && String(data.result).trim() !== "") {
    resultStr = String(data.result).trim();
  }

  sheet.getRange(data.row, 7).setValue(data.status || "Completed");
  sheet.getRange(data.row, 8).setValue(resultStr);

  recalculatePointsTable(data.tournamentId);
  return "Score saved (" + resultStr + ") & Points Table automatically updated!";
}

function recalculatePointsTable(tournamentId) {
  if (!tournamentId) return "Error: No tournament selected.";
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Detect sport for this tournament
  var sport = "Volleyball";
  var tSheet = ss.getSheetByName("Tournaments");
  if (tSheet && tSheet.getLastRow() >= 2) {
    var tData = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 2).getValues();
    for (var t = 0; t < tData.length; t++) {
      if (String(tData[t][0]).trim() === String(tournamentId).trim()) {
        sport = String(tData[t][1] || "Volleyball").trim();
        break;
      }
    }
  }
  var sportLower = sport.toLowerCase();
  var isVolleyball = (sportLower === "volleyball");
  var winPoints = (sportLower === "handball") ? 2 : 3;

  // 2. Initialize stats map with all Approved teams for this tournament
  var statsMap = {};
  var regSheet = ss.getSheetByName("Registration");
  if (regSheet && regSheet.getLastRow() >= 2) {
    var rCols = Math.max(regSheet.getLastColumn(), 9);
    var regData = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, rCols).getValues();
    for (var r = 0; r < regData.length; r++) {
      if (!regData[r][1]) continue;
      if (!matchesTournament(regData[r][8], tournamentId)) continue;
      var status = String(regData[r][7] || "").trim().toLowerCase();
      if (status === "approved") {
        var tName = String(regData[r][1]).trim();
        var key = tName.toLowerCase();
        statsMap[key] = {
          name: tName,
          mPld: 0, mW: 0, mL: 0, mPts: 0,
          sW: 0, sL: 0, sDiff: 0,
          pW: 0, pL: 0, pDiff: 0,
          logo: String(regData[r][5] || "No Logo"),
          division: String(regData[r][6] || "Men's")
        };
      }
    }
  }

  function ensureTeam(name) {
    var clean = String(name || "").trim();
    var key = clean.toLowerCase();
    if (!statsMap[key]) {
      statsMap[key] = {
        name: clean,
        mPld: 0, mW: 0, mL: 0, mPts: 0,
        sW: 0, sL: 0, sDiff: 0,
        pW: 0, pL: 0, pDiff: 0,
        logo: "No Logo",
        division: "Men's"
      };
    }
    return statsMap[key];
  }

  // 3. Process all Completed fixtures for this tournament
  var fixSheet = ss.getSheetByName("Fixtures");
  if (fixSheet && fixSheet.getLastRow() >= 2) {
    var fCols = Math.max(fixSheet.getLastColumn(), 9);
    var fixData = fixSheet.getRange(2, 1, fixSheet.getLastRow() - 1, fCols).getValues();

    for (var i = 0; i < fixData.length; i++) {
      if (!matchesTournament(fixData[i][8], tournamentId)) continue;
      var matchStatus = String(fixData[i][6] || "").trim().toLowerCase();
      if (matchStatus !== "completed" && matchStatus !== "final") continue;

      var teamAName = String(fixData[i][3] || "").trim();
      var teamBName = String(fixData[i][4] || "").trim();
      var resText = String(fixData[i][7] || "").trim();
      if (!teamAName || !teamBName || !resText) continue;

      // Match main score at start of result e.g. "3 - 1" or "2:2"
      var mainMatch = resText.match(/^\s*(\d+)\s*[-:]\s*(\d+)/);
      if (!mainMatch) continue;

      var scoreA = parseInt(mainMatch[1], 10);
      var scoreB = parseInt(mainMatch[2], 10);

      var teamA = ensureTeam(teamAName);
      var teamB = ensureTeam(teamBName);

      teamA.mPld += 1;
      teamB.mPld += 1;

      teamA.sW += scoreA;
      teamA.sL += scoreB;
      teamB.sW += scoreB;
      teamB.sL += scoreA;

      if (isVolleyball) {
        if (scoreA > scoreB) {
          teamA.mW += 1;
          teamB.mL += 1;
          if (scoreA - scoreB === 1) {
            teamA.mPts += 2;
            teamB.mPts += 1;
          } else {
            teamA.mPts += 3;
          }
        } else if (scoreB > scoreA) {
          teamB.mW += 1;
          teamA.mL += 1;
          if (scoreB - scoreA === 1) {
            teamB.mPts += 2;
            teamA.mPts += 1;
          } else {
            teamB.mPts += 3;
          }
        }

        // Optional set scores inside parentheses e.g. (25-20, 22-25, 25-18, 25-21)
        var parenMatch = resText.match(/\(([^)]+)\)/);
        if (parenMatch && parenMatch[1]) {
          var setRegex = /(\d+)\s*[-:]\s*(\d+)/g;
          var m;
          while ((m = setRegex.exec(parenMatch[1])) !== null) {
            var pA = parseInt(m[1], 10);
            var pB = parseInt(m[2], 10);
            teamA.pW += pA;
            teamA.pL += pB;
            teamB.pW += pB;
            teamB.pL += pA;
          }
        }
      } else {
        // Futsal & Handball (Draws stored in pW column)
        if (scoreA > scoreB) {
          teamA.mW += 1;
          teamA.mPts += winPoints;
          teamB.mL += 1;
        } else if (scoreB > scoreA) {
          teamB.mW += 1;
          teamB.mPts += winPoints;
          teamA.mL += 1;
        } else {
          teamA.pW += 1; // Draw
          teamB.pW += 1; // Draw
          teamA.mPts += 1;
          teamB.mPts += 1;
        }
      }
    }
  }

  // 4. Rebuild Points Table rows for this tournament (keeping other tournaments intact)
  var ptsSheet = ss.getSheetByName("Points Table");
  if (!ptsSheet) {
    ptsSheet = ss.insertSheet("Points Table");
    ptsSheet.appendRow(["Rank", "Team", "PLD", "W", "L", "PTS", "SW/GF", "SL/GA", "SDiff/GD", "PW/D", "PL", "PDiff", "Logo", "Division", "Tournament ID"]);
  }

  var keptRows = [];
  var lastPtsRow = ptsSheet.getLastRow();
  if (lastPtsRow >= 2) {
    var pCols = Math.max(ptsSheet.getLastColumn(), 15);
    var existingPts = ptsSheet.getRange(2, 1, lastPtsRow - 1, pCols).getValues();
    for (var j = 0; j < existingPts.length; j++) {
      if (!existingPts[j][1]) continue;
      if (!matchesTournament(existingPts[j][14], tournamentId)) {
        keptRows.push(existingPts[j].slice(0, 15));
      }
    }
  }

  var newRows = [];
  Object.keys(statsMap).forEach(function(k) {
    var s = statsMap[k];
    s.sDiff = s.sW - s.sL;
    s.pDiff = isVolleyball ? (s.pW - s.pL) : 0;
    newRows.push([
      "-", s.name, s.mPld, s.mW, s.mL, s.mPts,
      s.sW, s.sL, s.sDiff,
      s.pW, s.pL, s.pDiff,
      s.logo, s.division, tournamentId
    ]);
  });

  var finalRows = keptRows.concat(newRows);
  if (lastPtsRow >= 2) {
    ptsSheet.getRange(2, 1, lastPtsRow - 1, 15).clearContent();
  }
  if (finalRows.length > 0) {
    ptsSheet.getRange(2, 1, finalRows.length, 15).setValues(finalRows);
  }

  return "Points Table recalculated automatically for " + newRows.length + " teams!";
}

function updateTeamLogo(data) {
  if (!data) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  if (!sheet) return "Error: Registration sheet not found.";
  
  var numCols = Math.max(sheet.getLastColumn(), 9);
  var regData = sheet.getRange(1, 1, sheet.getLastRow(), numCols).getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < regData.length; i++) {
    if (matchesTournament(regData[i][8], data.tournamentId) && String(regData[i][1]).trim() === String(data.teamName).trim()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return "Error: Could not find team " + data.teamName + " in this tournament.";
  
  var logoUrl = "";
  if (data.logoLink && data.logoLink.trim() !== "") {
    logoUrl = data.logoLink.trim();
  } else if (data.logoBase64) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(data.logoBase64), data.logoMimeType, data.logoName);
      var file = DriveApp.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
      logoUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    } catch (e) {
      return "Error uploading image: " + e.message;
    }
  } else {
    return "Error: Please paste an image link or choose a file.";
  }
  
  sheet.getRange(rowIndex, 6).setValue(logoUrl);
  
  var ptsSheet = ss.getSheetByName("Points Table");
  if (ptsSheet && ptsSheet.getLastRow() >= 2) {
    var pCols = Math.max(ptsSheet.getLastColumn(), 15);
    var ptsData = ptsSheet.getRange(2, 1, ptsSheet.getLastRow() - 1, pCols).getValues();
    for (var j = 0; j < ptsData.length; j++) {
      if (matchesTournament(ptsData[j][14], data.tournamentId) && String(ptsData[j][1]).trim() === String(data.teamName).trim()) {
        ptsSheet.getRange(j + 2, 13).setValue(logoUrl);
        break;
      }
    }
  }
  
  return "Logo successfully updated for " + data.teamName + "!";
}