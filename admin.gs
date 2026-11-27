var SHEET_ID = '1Qv_TxvPY0_64ZVbVG2s8ZLvLyFJxtctCKsjealm23fE';
var DEFAULT_LEGACY_TOURNAMENT_ID = 'VB-2026';
var ADMIN_PASSWORD = 'admin123'; // <-- Change this to your preferred Admin password

function verifyAdminPassword(inputPassword) {
  return String(inputPassword || "").trim() === String(ADMIN_PASSWORD).trim();
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    try {
      var action = e.parameter.action;
      var tid = e.parameter.tournamentId || "";
      if (action === "getTournaments") return jsonResponse({ ok: true, data: getTournaments() });
      if (action === "getTeamsForApproval") return jsonResponse({ ok: true, data: getTeamsForApproval(tid) });
      if (action === "getTeams") return jsonResponse({ ok: true, data: getTeams(tid) });
      if (action === "getFixturesForAdmin") return jsonResponse({ ok: true, data: getFixturesForAdmin(tid) });
      return jsonResponse({ ok: false, error: "Unknown action: " + action });
    } catch (err) {
      return jsonResponse({ ok: false, error: err.message });
    }
  }

  var possibleNames = ['AdminIndex', 'Index', 'Admin', 'index', 'admin'];
  for (var i = 0; i < possibleNames.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(possibleNames[i])
          .setTitle('Admin Panel - Multi-Sport Tournament Portal')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) {
      // Try next file name
    }
  }
  throw new Error("No HTML file found. Please rename your HTML file in Apps Script to 'AdminIndex' or 'Index'.");
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var payload = body.payload || {};

    if (action === "verifyAdminPassword") return jsonResponse({ ok: true, data: verifyAdminPassword(payload.password) });
    if (action === "createTournament") return jsonResponse({ ok: true, data: createTournament(payload) });
    if (action === "updateTournamentStatus") return jsonResponse({ ok: true, data: updateTournamentStatus(payload.row, payload.status) });
    if (action === "updateTournamentSettings") return jsonResponse({ ok: true, data: updateTournamentSettings(payload) });
    if (action === "updateTeamStatus") return jsonResponse({ ok: true, data: updateTeamStatus(payload.row, payload.status, payload.group) });
    if (action === "updatePointsTable") return jsonResponse({ ok: true, data: updatePointsTable(payload) });
    if (action === "addFixture") return jsonResponse({ ok: true, data: addFixture(payload) });
    if (action === "updateMatchScore") return jsonResponse({ ok: true, data: updateMatchScore(payload) });
    if (action === "recalculatePointsTable") return jsonResponse({ ok: true, data: recalculatePointsTable(payload.tournamentId) });
    if (action === "updateTeamLogo") return jsonResponse({ ok: true, data: updateTeamLogo(payload) });

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function matchesTournament(cellValue, targetId) {
  var val = String(cellValue || "").trim();
  var target = String(targetId || "").trim();
  if (val === target) return true;
  if (val === "" && target === DEFAULT_LEGACY_TOURNAMENT_ID) return true;
  return false;
}

/**
 * Helper to save an uploaded base64 image to Google Drive or use a pasted link.
 */
function saveImageToDriveOrLink(logoLink, logoBase64, logoMimeType, logoName) {
  if (logoLink && String(logoLink).trim() !== "") {
    return String(logoLink).trim();
  }
  if (logoBase64) {
    try {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(logoBase64),
        logoMimeType || "image/png",
        logoName || "logo.png"
      );
      var file = DriveApp.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      return "https://lh3.googleusercontent.com/d/" + file.getId();
    } catch (e) {
      return "Error uploading: " + e.message;
    }
  }
  return "";
}

function getTournaments() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  if (!sheet) {
    sheet = ss.insertSheet("Tournaments");
    sheet.appendRow([
      "Tournament ID", "Sport", "Year", "Tournament Name", "Status",
      "Deadline", "Logo URL", "Rules", "Max Players", "Max Officials", "Divisions"
    ]);
    sheet.getRange("A1:K1").setFontWeight("bold");
    sheet.appendRow([
      DEFAULT_LEGACY_TOURNAMENT_ID, "Volleyball", 2026,
      "RCI Volleyball Tournament 2026", "Active", "", "", "", 14, 3, "Both"
    ]);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var numCols = Math.max(sheet.getLastColumn(), 11);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var tz = Session.getScriptTimeZone();
  var list = [];

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var deadlineStr = "";
    if (data[i][5] instanceof Date) {
      deadlineStr = Utilities.formatDate(data[i][5], tz, "yyyy-MM-dd'T'HH:mm");
    } else if (data[i][5]) {
      deadlineStr = String(data[i][5]).trim();
    }
    list.push({
      row: i + 2,
      id: String(data[i][0]).trim(),
      sport: String(data[i][1] || "Volleyball").trim(),
      year: String(data[i][2] || "").trim(),
      name: String(data[i][3] || "").trim(),
      status: String(data[i][4] || "Active").trim(),
      deadline: deadlineStr,
      logo: String(data[i][6] || "").trim(),
      rules: String(data[i][7] || ""),
      maxPlayers: data[i][8] || 14,
      maxOfficials: data[i][9] || 3,
      divisions: String(data[i][10] || "Both").trim()
    });
  }

  list.sort(function(a, b) {
    return (parseInt(b.year, 10) || 0) - (parseInt(a.year, 10) || 0);
  });
  return list;
}

function createTournament(data) {
  if (!data) return "Error: Missing tournament data.";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  if (!sheet) {
    getTournaments();
    sheet = ss.getSheetByName("Tournaments");
  }

  var prefix = "TRN";
  var sLower = String(data.sport || "").toLowerCase();
  if (sLower === "volleyball") prefix = "VB";
  else if (sLower === "futsal") prefix = "FS";
  else if (sLower === "handball") prefix = "HB";

  var newId = prefix + "-" + data.year + "-" + Math.floor(100 + Math.random() * 900);
  var logoUrl = saveImageToDriveOrLink(data.logoLink, data.logoBase64, data.logoMimeType, data.logoName);

  sheet.appendRow([
    newId,
    data.sport,
    data.year,
    String(data.name || "").trim(),
    data.status || "Active",
    data.deadline || "",
    logoUrl,
    data.rules || "",
    data.maxPlayers || 14,
    data.maxOfficials || 3,
    data.divisions || "Both"
  ]);

  return "New tournament '" + data.name + "' (" + data.year + ") created successfully!";
}

function updateTournamentStatus(row, status) {
  if (!row || !status) return "Error: Invalid tournament row or status.";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  if (!sheet) return "Error: Tournaments sheet not found.";
  sheet.getRange(row, 5).setValue(status);
  return "Tournament status updated to " + status + "!";
}

function updateTournamentSettings(data) {
  if (!data || !data.tournamentId) return "Error: Missing tournament ID.";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Tournaments");
  if (!sheet) return "Error: Tournaments sheet not found.";

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "Error: Tournament not found.";
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var rowIndex = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(data.tournamentId).trim()) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) return "Error: Tournament ID not found.";

  // Ensure Column K header exists on older sheets
  if (!sheet.getRange(1, 11).getValue()) {
    sheet.getRange(1, 11).setValue("Divisions");
  }

  sheet.getRange(rowIndex, 4).setValue(String(data.name || "").trim());
  sheet.getRange(rowIndex, 5).setValue(data.status || "Active");
  sheet.getRange(rowIndex, 6).setValue(data.deadline || "");
  sheet.getRange(rowIndex, 8).setValue(data.rules || "");
  sheet.getRange(rowIndex, 9).setValue(data.maxPlayers || 14);
  sheet.getRange(rowIndex, 10).setValue(data.maxOfficials || 3);
  sheet.getRange(rowIndex, 11).setValue(data.divisions || "Both");

  if (data.logoBase64) {
    var uploadedLogo = saveImageToDriveOrLink("", data.logoBase64, data.logoMimeType, data.logoName);
    if (uploadedLogo && uploadedLogo.indexOf("Error") !== 0) {
      sheet.getRange(rowIndex, 7).setValue(uploadedLogo);
    }
  } else if (data.logoLink !== undefined) {
    sheet.getRange(rowIndex, 7).setValue(String(data.logoLink).trim());
  }

  SpreadsheetApp.flush();
  return "Tournament settings saved successfully!";
}

function getTeamsForApproval(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numCols = Math.max(sheet.getLastColumn(), 10);
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
      status: String(data[i][7] || "Pending"),
      group: String(data[i][9] || "").trim()
    });
  }
  return teams;
}

function updateTeamStatus(row, status, group) {
  if (!row || !status) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  sheet.getRange(row, 8).setValue(status);
  if (group !== undefined) {
    sheet.getRange(row, 10).setValue(group);
    // Sync group to Points Table if team already exists there
    var teamName = String(sheet.getRange(row, 2).getValue() || "").trim();
    var tournId = String(sheet.getRange(row, 9).getValue() || "").trim();
    var ptsSheet = ss.getSheetByName("Points Table");
    if (ptsSheet && ptsSheet.getLastRow() >= 2 && teamName) {
      var pCols = Math.max(ptsSheet.getLastColumn(), 16);
      var ptsData = ptsSheet.getRange(2, 1, ptsSheet.getLastRow() - 1, pCols).getValues();
      for (var j = 0; j < ptsData.length; j++) {
        if (matchesTournament(ptsData[j][14], tournId || DEFAULT_LEGACY_TOURNAMENT_ID) && String(ptsData[j][1]).trim() === teamName) {
          ptsSheet.getRange(j + 2, 16).setValue(group);
          break;
        }
      }
    }
  }
  return "Team updated (" + status + (group ? " | " + group : "") + ")";
}

function getTeams(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numCols = Math.max(sheet.getLastColumn(), 10);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var teams = [];

  for (var i = 0; i < data.length; i++) {
    if (!data[i][1]) continue;
    if (!matchesTournament(data[i][8], tournamentId)) continue;
    var status = String(data[i][7] || "").trim().toLowerCase();
    if (status === "approved") {
      teams.push(String(data[i][1]).trim());
    }
  }
  return teams;
}

function updatePointsTable(data) {
  if (!data) return "Success: Script authorized!";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Points Table");
  if (!sheet) {
    sheet = ss.insertSheet("Points Table");
    sheet.appendRow(["Rank", "Team", "PLD", "W", "L", "PTS", "SW/GF", "SL/GA", "SDiff/GD", "PW/D", "PL", "PDiff", "Logo", "Division", "Tournament ID", "Group"]);
  }

  var regSheet = ss.getSheetByName("Registration");
  var logoUrl = "No Logo";
  var division = "Men's";
  var group = "";
  if (regSheet && regSheet.getLastRow() >= 2) {
    var rCols = Math.max(regSheet.getLastColumn(), 10);
    var regData = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, rCols).getValues();
    for (var j = 0; j < regData.length; j++) {
      if (matchesTournament(regData[j][8], data.tournamentId) && String(regData[j][1]).trim() === String(data.teamName).trim()) {
        logoUrl = regData[j][5];
        division = regData[j][6];
        group = String(regData[j][9] || "").trim();
        break;
      }
    }
  }

  var rowData = [
    "-", data.teamName, data.mPld, data.mW, data.mL, data.mPts,
    data.sW, data.sL, data.sDiff,
    data.pW, data.pL, data.pDiff,
    logoUrl, division, data.tournamentId, group
  ];

  var lastRow = sheet.getLastRow();
  var rowIndex = -1;
  if (lastRow >= 2) {
    var numCols = Math.max(sheet.getLastColumn(), 16);
    var dataRange = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    for (var i = 0; i < dataRange.length; i++) {
      if (matchesTournament(dataRange[i][14], data.tournamentId) && String(dataRange[i][1]).trim() === String(data.teamName).trim()) {
        rowIndex = i + 2;
        break;
      }
    }
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 16).setValues([rowData]);
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
    sheet.appendRow(["Match ID", "Date", "Time", "Team A", "Team B", "Venue", "Status", "Result", "Tournament ID", "Group"]);
  }
  sheet.appendRow([
    data.matchId, data.date, data.time,
    data.teamA, data.teamB, data.venue,
    data.status, data.result, data.tournamentId,
    data.group || ""
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

  var numCols = Math.max(sheet.getLastColumn(), 10);
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
      result: String(data[i][7] || ""),
      group: String(data[i][9] || "").trim()
    });
  }
  return list;
}

function updateMatchScore(data) {
  if (!data || !data.row) return "Error: Invalid match row.";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Fixtures");
  if (!sheet) return "Error: Fixtures sheet not found.";

  var resultStr = "-";
  if (data.scoreA !== "" && data.scoreB !== "") {
    resultStr = String(data.scoreA).trim() + " - " + String(data.scoreB).trim();
    if (data.setDetails && String(data.setDetails).trim() !== "") {
      resultStr += " (" + String(data.setDetails).trim() + ")";
    }
  }

  sheet.getRange(data.row, 7).setValue(data.status || "Completed");
  sheet.getRange(data.row, 8).setValue(resultStr);

  recalculatePointsTable(data.tournamentId);
  return "Match score saved & Points Table automatically updated!";
}

function recalculatePointsTable(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Determine sport type for this tournament
  var sport = "Volleyball";
  var allTourns = getTournaments();
  for (var t = 0; t < allTourns.length; t++) {
    if (allTourns[t].id === String(tournamentId).trim()) {
      sport = allTourns[t].sport;
      break;
    }
  }
  var sportLower = sport.toLowerCase();
  var isVolleyball = (sportLower === "volleyball");
  var isHandball = (sportLower === "handball");

  // 2. Initialize stats map with all Approved teams for this tournament
  var statsMap = {};
  var regSheet = ss.getSheetByName("Registration");
  if (regSheet && regSheet.getLastRow() >= 2) {
    var rCols = Math.max(regSheet.getLastColumn(), 10);
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
          division: String(regData[r][6] || "Men's"),
          group: String(regData[r][9] || "").trim()
        };
      }
    }
  }

  function ensureTeam(name, fallbackGroup) {
    var clean = String(name || "").trim();
    var key = clean.toLowerCase();
    if (!statsMap[key]) {
      statsMap[key] = {
        name: clean,
        mPld: 0, mW: 0, mL: 0, mPts: 0,
        sW: 0, sL: 0, sDiff: 0,
        pW: 0, pL: 0, pDiff: 0,
        logo: "No Logo",
        division: "Men's",
        group: fallbackGroup || ""
      };
    } else if (!statsMap[key].group && fallbackGroup) {
      statsMap[key].group = fallbackGroup;
    }
    return statsMap[key];
  }

  // 3. Process all Completed group/league fixtures for this tournament
  var fixSheet = ss.getSheetByName("Fixtures");
  if (fixSheet && fixSheet.getLastRow() >= 2) {
    var fCols = Math.max(fixSheet.getLastColumn(), 10);
    var fixData = fixSheet.getRange(2, 1, fixSheet.getLastRow() - 1, fCols).getValues();

    for (var i = 0; i < fixData.length; i++) {
      if (!matchesTournament(fixData[i][8], tournamentId)) continue;
      var matchStatus = String(fixData[i][6] || "").trim().toLowerCase();
      if (matchStatus !== "completed" && matchStatus !== "final") continue;

      // Skip knockout matches from group standings
      var matchGroup = String(fixData[i][9] || "").trim();
      var mgLower = matchGroup.toLowerCase();
      if (mgLower.indexOf("quarter") > -1 || mgLower.indexOf("semi") > -1 || mgLower.indexOf("final") > -1 || mgLower.indexOf("3rd") > -1 || mgLower.indexOf("knockout") > -1) {
        continue;
      }

      var teamAName = String(fixData[i][3] || "").trim();
      var teamBName = String(fixData[i][4] || "").trim();
      var resText = String(fixData[i][7] || "").trim();
      if (!teamAName || !teamBName || !resText) continue;

      var mainMatch = resText.match(/^\s*(\d+)\s*[-:]\s*(\d+)/);
      if (!mainMatch) continue;

      var scoreA = parseInt(mainMatch[1], 10);
      var scoreB = parseInt(mainMatch[2], 10);

      var teamA = ensureTeam(teamAName, matchGroup);
      var teamB = ensureTeam(teamBName, matchGroup);

      teamA.mPld += 1;
      teamB.mPld += 1;
      teamA.sW += scoreA;
      teamA.sL += scoreB;
      teamB.sW += scoreB;
      teamB.sL += scoreA;

      if (isVolleyball) {
        // Volleyball scoring: 3-0 / 3-1 => 3/0 pts; 3-2 / 2-1 => 2/1 pts
        if (scoreA > scoreB) {
          teamA.mW += 1;
          teamB.mL += 1;
          if (scoreA - scoreB === 1 && scoreB >= 1) {
            teamA.mPts += 2;
            teamB.mPts += 1;
          } else {
            teamA.mPts += 3;
          }
        } else if (scoreB > scoreA) {
          teamB.mW += 1;
          teamA.mL += 1;
          if (scoreB - scoreA === 1 && scoreA >= 1) {
            teamB.mPts += 2;
            teamA.mPts += 1;
          } else {
            teamB.mPts += 3;
          }
        }
        // Parse optional set points inside parentheses e.g. "(25-20, 25-18, 25-22)"
        var parenMatch = resText.match(/\(([^)]+)\)/);
        if (parenMatch && parenMatch[1]) {
          var setPairs = parenMatch[1].split(",");
          setPairs.forEach(function(pair) {
            var sp = pair.match(/(\d+)\s*[-:]\s*(\d+)/);
            if (sp) {
              var pA = parseInt(sp[1], 10);
              var pB = parseInt(sp[2], 10);
              teamA.pW += pA;
              teamA.pL += pB;
              teamB.pW += pB;
              teamB.pL += pA;
            }
          });
        }
      } else {
        // Futsal (W=3, D=1, L=0) & Handball (W=2, D=1, L=0); Draws stored in pW column
        var winPts = isHandball ? 2 : 3;
        if (scoreA > scoreB) {
          teamA.mW += 1;
          teamB.mL += 1;
          teamA.mPts += winPts;
        } else if (scoreB > scoreA) {
          teamB.mW += 1;
          teamA.mL += 1;
          teamB.mPts += winPts;
        } else {
          teamA.pW += 1;
          teamB.pW += 1;
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
    ptsSheet.appendRow(["Rank", "Team", "PLD", "W", "L", "PTS", "SW/GF", "SL/GA", "SDiff/GD", "PW/D", "PL", "PDiff", "Logo", "Division", "Tournament ID", "Group"]);
  }

  var keptRows = [];
  var lastPtsRow = ptsSheet.getLastRow();
  if (lastPtsRow >= 2) {
    var pCols = Math.max(ptsSheet.getLastColumn(), 16);
    var existingPts = ptsSheet.getRange(2, 1, lastPtsRow - 1, pCols).getValues();
    for (var j = 0; j < existingPts.length; j++) {
      if (!existingPts[j][1]) continue;
      if (!matchesTournament(existingPts[j][14], tournamentId)) {
        keptRows.push(existingPts[j].slice(0, 16));
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
      s.logo, s.division, tournamentId, s.group || ""
    ]);
  });

  var finalRows = keptRows.concat(newRows);
  if (lastPtsRow >= 2) {
    ptsSheet.getRange(2, 1, lastPtsRow - 1, 16).clearContent();
  }
  if (finalRows.length > 0) {
    ptsSheet.getRange(2, 1, finalRows.length, 16).setValues(finalRows);
  }

  return "Points Table recalculated automatically for " + newRows.length + " teams!";
}

function updateTeamLogo(data) {
  if (!data || !data.teamName) return "Error: Missing team name.";
  var logoUrl = saveImageToDriveOrLink(data.logoLink, data.logoBase64, data.logoMimeType, data.logoName);
  if (!logoUrl || logoUrl.indexOf("Error") === 0) {
    return logoUrl || "Error: Please provide a valid image link or file.";
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var targetName = String(data.teamName).trim();

  var regSheet = ss.getSheetByName("Registration");
  if (regSheet && regSheet.getLastRow() >= 2) {
    var rCols = Math.max(regSheet.getLastColumn(), 9);
    var regData = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, rCols).getValues();
    for (var i = 0; i < regData.length; i++) {
      if (matchesTournament(regData[i][8], data.tournamentId) && String(regData[i][1]).trim() === targetName) {
        regSheet.getRange(i + 2, 6).setValue(logoUrl);
        break;
      }
    }
  }

  var ptsSheet = ss.getSheetByName("Points Table");
  if (ptsSheet && ptsSheet.getLastRow() >= 2) {
    var pCols = Math.max(ptsSheet.getLastColumn(), 15);
    var ptsData = ptsSheet.getRange(2, 1, ptsSheet.getLastRow() - 1, pCols).getValues();
    for (var j = 0; j < ptsData.length; j++) {
      if (matchesTournament(ptsData[j][14], data.tournamentId) && String(ptsData[j][1]).trim() === targetName) {
        ptsSheet.getRange(j + 2, 13).setValue(logoUrl);
        break;
      }
    }
  }

  return "Logo updated successfully for " + targetName + "!";
}
