var SHEET_ID = '1Qv_TxvPY0_64ZVbVG2s8ZLvLyFJxtctCKsjealm23fE';
var DEFAULT_LEGACY_TOURNAMENT_ID = 'VB-2026';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sports Tournament Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Ensures the "Tournaments" sheet exists and returns all tournaments.
 * Any old data without a Tournament ID will automatically link to DEFAULT_LEGACY_TOURNAMENT_ID.
 */
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
    
    // Migrate existing settings if available
    var oldSettings = ss.getSheetByName("Settings");
    var name = "RCI Volleyball Tournament 2026";
    var deadline = "";
    var logo = "";
    var rules = "";
    var maxP = 14;
    var maxO = 3;
    if (oldSettings) {
      if (oldSettings.getRange("B1").getValue()) name = oldSettings.getRange("B1").getValue();
      if (oldSettings.getRange("B2").getValue()) deadline = oldSettings.getRange("B2").getValue();
      if (oldSettings.getRange("B3").getValue()) logo = oldSettings.getRange("B3").getValue();
      if (oldSettings.getRange("B4").getValue()) rules = oldSettings.getRange("B4").getValue();
      if (oldSettings.getRange("B5").getValue()) maxP = oldSettings.getRange("B5").getValue();
      if (oldSettings.getRange("B6").getValue()) maxO = oldSettings.getRange("B6").getValue();
    }
    sheet.appendRow([DEFAULT_LEGACY_TOURNAMENT_ID, "Volleyball", 2026, name, "Active", deadline, logo, rules, maxP, maxO]);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var now = new Date();
  var list = [];
  
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rawDeadline = data[i][5];
    var status = String(data[i][4] || "Active").trim();
    var isClosed = (status.toLowerCase() === "archived" || status.toLowerCase() === "closed");
    
    if (!isClosed && rawDeadline) {
      var dDate = new Date(rawDeadline);
      if (!isNaN(dDate.getTime()) && now > dDate) {
        isClosed = true;
      }
    }
    
    list.push({
      id: String(data[i][0]).trim(),
      sport: String(data[i][1] || "Volleyball").trim(),
      year: String(data[i][2] || "").trim(),
      name: String(data[i][3] || "Tournament").trim(),
      status: status,
      deadline: rawDeadline ? String(rawDeadline) : "",
      logo: String(data[i][6] || "").trim(),
      rules: String(data[i][7] || ""),
      maxPlayers: parseInt(data[i][8]) || 14,
      maxOfficials: parseInt(data[i][9]) || 3,
      isClosed: isClosed
    });
  }
  
  // Sort newest year first
  list.sort(function(a, b) {
    return (parseInt(b.year) || 0) - (parseInt(a.year) || 0);
  });
  
  return list;
}

function getTournamentById(tournamentId) {
  var all = getTournaments();
  for (var i = 0; i < all.length; i++) {
    if (all[i].id === String(tournamentId).trim()) return all[i];
  }
  return all.length > 0 ? all[0] : null;
}

function matchesTournament(cellValue, targetId) {
  var val = String(cellValue || "").trim();
  var target = String(targetId || "").trim();
  if (val === target) return true;
  // Backward compatibility for legacy rows with blank Tournament ID
  if (val === "" && target === DEFAULT_LEGACY_TOURNAMENT_ID) return true;
  return false;
}

function registerTeam(data) {
  var tourn = getTournamentById(data.tournamentId);
  if (!tourn) return "Error: Selected tournament not found.";
  if (tourn.isClosed) return "Error: Registration for " + tourn.name + " is closed.";

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  if (!sheet) {
    sheet = ss.insertSheet("Registration");
    sheet.appendRow(["Timestamp", "Team Name", "Captain", "Email", "Phone", "Logo URL", "Division", "Status", "Tournament ID"]);
  }
  
  var existingData = sheet.getDataRange().getValues();
  var newTeamName = data.teamName.trim().toLowerCase();
  
  // Check for duplicate team name within the SAME tournament
  for (var i = 1; i < existingData.length; i++) {
    if (matchesTournament(existingData[i][8], tourn.id)) {
      if (existingData[i][1] && existingData[i][1].toString().trim().toLowerCase() === newTeamName) {
        return "Error: The team name '" + data.teamName + "' is already registered in this tournament.";
      }
    }
  }

  var logoUrl = "No Logo";
  if (data.logoLink && data.logoLink.trim() !== "") {
    logoUrl = data.logoLink.trim();
  } else if (data.logoBase64) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(data.logoBase64), data.logoMimeType, data.logoName);
      var file = DriveApp.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      logoUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    } catch (e) {
      logoUrl = "Error uploading";
    }
  }
  
  sheet.appendRow([
    new Date(),
    data.teamName.trim(),
    data.captainName.trim(),
    data.email.trim(),
    data.phone.trim(),
    logoUrl,
    data.division,
    "Pending",
    tourn.id
  ]);
  return "Team registered successfully for " + tourn.name + "! Awaiting Admin Approval.";
}

function registerPlayers(data) {
  var tourn = getTournamentById(data.tournamentId);
  if (!tourn) return "Error: Selected tournament not found.";
  if (tourn.isClosed) return "Error: Player submission for " + tourn.name + " is closed.";

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var regSheet = ss.getSheetByName("Registration");
  if (!regSheet) return "Error: Registration sheet not found.";
  
  var regData = regSheet.getDataRange().getValues();
  var isApproved = false;
  var targetTeam = String(data.teamName).trim().toLowerCase();
  
  for (var i = 1; i < regData.length; i++) {
    if (matchesTournament(regData[i][8], tourn.id)) {
      var rowTeam = String(regData[i][1] || "").trim().toLowerCase();
      var rowStatus = String(regData[i][7] || "").trim().toLowerCase();
      if (rowTeam === targetTeam && rowStatus === "approved") {
        isApproved = true;
        break;
      }
    }
  }
  
  if (!isApproved) return "Error: Team must be approved by an Admin before submitting players.";

  var sheet = ss.getSheetByName("Players");
  if (!sheet) {
    sheet = ss.insertSheet("Players");
    sheet.appendRow(["Timestamp", "Team Name", "Full Name", "Jersey", "Position/Role", "Phone", "Tournament ID"]);
  }
  
  var timestamp = new Date();
  data.members.forEach(function(member) {
    sheet.appendRow([
      timestamp,
      data.teamName.trim(),
      member.name.trim(),
      member.jerseyNumber,
      member.position,
      member.phone,
      tourn.id
    ]);
  });
  
  return "Full roster submitted successfully for " + data.teamName + "!";
}

function getPointsTable(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Points Table");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var numCols = Math.max(sheet.getLastColumn(), 15);
  var raw = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var filtered = [];
  
  for (var i = 0; i < raw.length; i++) {
    if (!raw[i][1]) continue;
    if (matchesTournament(raw[i][14], tournamentId)) {
      filtered.push(raw[i].slice(0, 15));
    }
  }
  
  // Sort by PTS (col 5), then Set Diff / Goal Diff (col 8), then Pts Diff / Goals For (col 11 or 6)
  filtered.sort(function(a, b) {
    var ptsA = Number(a[5]) || 0, ptsB = Number(b[5]) || 0;
    if (ptsB !== ptsA) return ptsB - ptsA;
    var diffA = Number(a[8]) || 0, diffB = Number(b[8]) || 0;
    if (diffB !== diffA) return diffB - diffA;
    var tieA = Number(a[11]) || Number(a[6]) || 0;
    var tieB = Number(b[11]) || Number(b[6]) || 0;
    return tieB - tieA;
  });
  
  return filtered;
}

function getFixtures(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Fixtures");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var numCols = Math.max(sheet.getLastColumn(), 9);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var filtered = [];
  var tz = Session.getScriptTimeZone();
  
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0] && !data[i][3]) continue;
    if (!matchesTournament(data[i][8], tournamentId)) continue;
    
    var row = data[i].slice(0, 9);
    if (row[1] instanceof Date) {
      row[1] = Utilities.formatDate(row[1], tz, "yyyy-MM-dd");
    }
    if (row[2] instanceof Date) {
      row[2] = Utilities.formatDate(row[2], tz, "HH:mm");
    }
    filtered.push(row);
  }
  return filtered;
}

function getTeams(tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Registration");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var numCols = Math.max(sheet.getLastColumn(), 9);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var approvedTeams = [];
  
  for (var i = 0; i < data.length; i++) {
    if (!data[i][1]) continue;
    if (!matchesTournament(data[i][8], tournamentId)) continue;
    
    var status = String(data[i][7] || "").trim().toLowerCase();
    if (status === "approved") {
      approvedTeams.push({
        name: String(data[i][1]).trim(),
        logo: String(data[i][5] || "").trim(),
        division: String(data[i][6] || "Men's").trim()
      });
    }
  }
  return approvedTeams;
}

function getTeamRoster(teamName, tournamentId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Find "Players" or "Player" sheet even if there are extra spaces or different casing
  var sheet = ss.getSheetByName("Players") || ss.getSheetByName("Player");
  if (!sheet) {
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var sName = allSheets[s].getName().trim().toLowerCase();
      if (sName === "players" || sName === "player" || sName.indexOf("player") > -1) {
        sheet = allSheets[s];
        break;
      }
    }
  }
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var numCols = Math.max(sheet.getLastColumn(), 7);
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  
  // Keep letters from all languages (including Dhivehi), ignore spaces and punctuation
  function normalizeName(str) {
    return String(str || "").replace(/[\s'"`.,\-_()]+/g, "").toLowerCase();
  }
  
  var targetTeam = normalizeName(teamName);
  var targetTourn = String(tournamentId || "").trim();
  var exactTournamentRoster = [];
  var untaggedRoster = [];
  
  for (var i = 0; i < data.length; i++) {
    if (!data[i][1]) continue;
    
    var currentTeam = normalizeName(data[i][1]);
    if (currentTeam !== targetTeam) continue;
    
    var playerObj = {
      name: String(data[i][2] || "").trim(),
      jersey: String(data[i][3] || "").trim(),
      position: String(data[i][4] || "").trim(),
      phone: String(data[i][5] || "").trim()
    };
    
    var rowTourn = String(data[i][6] || "").trim();
    if (rowTourn !== "" && rowTourn === targetTourn) {
      exactTournamentRoster.push(playerObj);
    } else if (rowTourn === "") {
      untaggedRoster.push(playerObj);
    }
  }
  
  // If players tagged with this specific tournament exist, return them; otherwise fallback to legacy untagged rows
  return exactTournamentRoster.length > 0 ? exactTournamentRoster : untaggedRoster;
}