// ============================================================
// Fast Facts — StrawPoll Automated Leaderboard (Direct API Pull)
// ============================================================

const scriptProps = PropertiesService.getScriptProperties();

const CONFIG = {
  // Credentials pull securely from Apps Script > Project Settings > Script Properties
  API_KEY:         scriptProps.getProperty('STRAWPOLL_API_KEY') || 'YOUR_STRAWPOLL_API_KEY_HERE',
  SHEET_ID:        scriptProps.getProperty('SHEET_ID') || 'YOUR_GOOGLE_SHEET_ID_HERE',
  
  LEADERBOARD_TAB: '2026 Scores',   
  CONFIG_TAB:      'Config',
  POLL_LOG_TAB:    'Poll Log',
  SCRIPT_LOG_TAB:  'Script Log',
  ALIASES_TAB:     'Aliases',
  SNAPSHOTS_TAB:   'Snapshots',        

  DATA_START_ROW: 1,

  // Columns in the leaderboard tab (1-indexed)
  COL_RANK:  1,   // A — numeric rank
  COL_NAME:  3,   // C — player name
  COL_SCORE: 4,   // D — total score

  PDF_RANGE:       'B1:D12',          
  PDF_FOLDER_NAME: 'Fast Facts PDFs', 
};


// ============================================================
// BULLETPROOF WEBHOOK HANDLER
// ============================================================

function doPost(e) {
  let rawPayload = '(no body)';
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const logSh = ss.getSheetByName(CONFIG.SCRIPT_LOG_TAB);
    const ts    = new Date().toISOString();

    rawPayload = e && e.postData && e.postData.contents ? e.postData.contents : '(no body)';
    const payload = JSON.parse(rawPayload);
    
    const eventType = String(payload.event || payload.type || '').toLowerCase();
    
    // Extract poll ID across all potential StrawPoll payload variations
    const pollId = (payload.data && payload.data.poll && payload.data.poll.id) || 
                   (payload.data && payload.data.id) || 
                   payload.poll_id || 
                   payload.id || 
                   '?';

    if (logSh) {
      logSh.appendRow([ts, 'Webhook ping | Event: ' + eventType + ' | Poll ID: ' + pollId]);
    }

    // Broad pattern match: catches deadline, closed, update, or status flags
    const isClosureEvent = eventType.includes('closed') || 
                           eventType.includes('deadline') || 
                           eventType.includes('update') || 
                           eventType.includes('end') ||
                           (payload.data && (payload.data.status === 'closed' || payload.data.is_closed === true));

    if (isClosureEvent && pollId !== '?') {
      if (logSh) logSh.appendRow([ts, '⚡ Closure event verified for poll ' + pollId + '. Executing forced scoring...']);
      
      // Pass pollId directly to bypass status API checks
      checkPolls(pollId);

      return ContentService.createTextOutput('SUCCESS: Scoring triggered for poll ' + pollId).setMimeType(ContentService.MimeType.TEXT);
    }

    return ContentService.createTextOutput('ACK: Event received but ignored (' + eventType + ')').setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput('APPS SCRIPT ERROR: ' + err.message + ' | Payload: ' + rawPayload).setMimeType(ContentService.MimeType.TEXT);
  }
}


// ============================================================
// MAIN SCORING ENGINE
// ============================================================

function checkPolls(forcePollId) {
  const ss       = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const configSh = getOrCreateSheet(ss, CONFIG.CONFIG_TAB, ['Poll ID', 'Week Label', 'Correct Answer', 'Processed', 'Processed At']);
  const logSh    = getOrCreateSheet(ss, CONFIG.SCRIPT_LOG_TAB, ['Timestamp', 'Message']);

  log(logSh, '— checkPolls execution started —');
  const rows = configSh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const [pollId, weekLabel, correctAnswer, processed] = rows[i];

    if (!pollId || processed === true || processed === 'TRUE') continue;
    if (!correctAnswer || !String(correctAnswer).trim()) {
      log(logSh, 'Poll ' + pollId + ': skipping — correct answer missing in Config tab.');
      continue;
    }

    log(logSh, 'Processing poll: ' + pollId + ' (' + weekLabel + ')');

    try {
      const pollData = fetchPollResultsData(pollId, logSh);

      // Verify if scoring was explicitly requested by a webhook trigger
      const normalizedForce = forcePollId ? String(forcePollId).trim().toLowerCase() : '';
      const normalizedCurrent = String(pollId).trim().toLowerCase();
      const isForced = normalizedForce !== '' && normalizedForce === normalizedCurrent;

      if (!isForced) {
        if (!isPollClosed(pollData)) {
          log(logSh, 'Poll ' + pollId + ' appears active — skipping manual run.');
          continue;
        }
      } else {
        log(logSh, 'Poll ' + pollId + ' closure confirmed by webhook flag.');
      }

      const optionMap = buildOptionMap(pollData);
      const votes     = extractVotesFromResults(pollData);
      log(logSh, 'Votes retrieved from API: ' + votes.length);

      if (votes.length === 0) {
        log(logSh, 'WARNING: No votes found for poll ' + pollId);
        configSh.getRange(i + 1, 4).setValue('TRUE — NO PARTICIPANT DATA');
        continue;
      }

      const rankSnapshot = getLeaderboardSnapshot(ss);
      const aliases      = loadAliases(ss);
      const correctLabel = resolveCorrectAnswer(correctAnswer, optionMap);
      
      // Resolve aliases and deduplicate votes upfront
      const resolvedVotesMap = {};
      for (const vote of votes) {
        const canonicalName = fuzzyMatchName(vote.name, rankSnapshot, aliases) || vote.name;
        const key = normName(canonicalName);
        
        // Deduplicate: If they voted multiple times, keep only their earliest submission
        if (!resolvedVotesMap[key] || new Date(vote.timestamp) < new Date(resolvedVotesMap[key].timestamp)) {
          resolvedVotesMap[key] = { ...vote, originalName: vote.name, name: canonicalName };
        }
      }
      const processedVotes = Object.values(resolvedVotesMap);

      const firstSubmitter  = getFirstSubmitter(processedVotes);
      const correct         = correctLabel.toLowerCase();
      const firstGotItRight = firstSubmitter && firstSubmitter.answer.trim().toLowerCase() === correct;

      saveSnapshot(ss, pollId, weekLabel);
      
      // Both the logs and the scoring engine use the clean, canonical data
      writePollLog(ss, weekLabel, pollId, processedVotes, correctLabel, firstSubmitter, rankSnapshot);
      applyScores(ss, processedVotes, correct, firstSubmitter, firstGotItRight, rankSnapshot);

      configSh.getRange(i + 1, 4).setValue('TRUE');
      configSh.getRange(i + 1, 5).setValue(new Date().toISOString());

      // FORCE GOOGLE TO COMMIT ALL LEADERBOARD CELL CHANGES TO DISK BEFORE PDF GENERATION
      SpreadsheetApp.flush();

      try {
        const pdfUrl = exportLeaderboardPDF(ss, weekLabel);
        log(logSh, 'PDF exported successfully: ' + pdfUrl);
      } catch (pdfErr) {
        log(logSh, 'PDF export encountered non-fatal error: ' + pdfErr.message);
      }

      log(logSh, 'Poll ' + pollId + ' fully processed.');

    } catch (err) {
      log(logSh, 'ERROR processing poll ' + pollId + ': ' + err.message);
    }
  }
  log(logSh, '— checkPolls execution complete —');
}


// ============================================================
// STRAWPOLL INTEGRATION
// ============================================================

function fetchPollResultsData(pollId, logSh) {
  const url = 'https://api.strawpoll.com/v3/polls/' + pollId + '/results';
  const res = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + CONFIG.API_KEY },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('StrawPoll API returned HTTP ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 100));
  }

  const json = JSON.parse(res.getContentText());
  return json.data || json;
}

function extractVotesFromResults(pollData) {
  const participants = pollData.poll_participants || [];
  const options      = pollData.poll_options || [];
  if (!participants.length || !options.length) return [];

  const sortedOptions = [...options].sort((a, b) => (a.position || 0) - (b.position || 0));
  const votes = [];

  for (const p of participants) {
    const name = String(p.name || 'Unknown').trim();
    const ts   = p.created_at ? new Date(p.created_at * 1000).toISOString() : null;
    const pollVotes = p.poll_votes || [];

    for (let i = 0; i < pollVotes.length; i++) {
      if (pollVotes[i] === 1 && sortedOptions[i]) {
        const label = String(sortedOptions[i].value || sortedOptions[i].label || ('Option ' + (i + 1))).trim();
        votes.push({ name, answer: label, timestamp: ts });
      }
    }
  }
  return votes;
}

function isPollClosed(poll) {
  if (!poll) return false;

  if (poll.is_votable === 0 || poll.is_votable === false) return true;
  if (poll.is_closed === true || poll.closed_at) return true;

  const status = String(poll.status || poll.poll_status || '').toLowerCase();
  if (status === 'closed' || status === '2' || status === 'disabled' || status === 'ended' || status === 'expired') return true;

  const deadline = (poll.poll_config && poll.poll_config.deadline_at) || poll.deadline_at;
  if (deadline) {
    const deadlineMs = deadline > 1e11 ? deadline : deadline * 1000;
    if (Date.now() >= deadlineMs) return true;
  }

  if (status === 'live' || status === 'open' || status === 'active' || poll.is_votable === true) {
    return false;
  }

  return true; 
}

function buildOptionMap(poll) {
  const map = {};
  const options = poll.poll_options || poll.options || [];
  const sorted = [...options].sort((a, b) => (a.position || 0) - (b.position || 0));
  sorted.forEach((opt, idx) => {
    const label = String(opt.value || opt.label || opt.text || ('Option ' + opt.id)).trim();
    map[opt.id] = label;  
    map[String(idx + 1)] = label;  
  });
  return map;
}

function resolveCorrectAnswer(correctAnswerRaw, optionMap) {
  const raw = String(correctAnswerRaw).trim();
  if (/^[1-6]$/.test(raw) && optionMap[raw]) return optionMap[raw];
  return raw;
}


// ============================================================
// CALCULATIONS & SCORE PROCESSING
// ============================================================

function getLeaderboardSnapshot(ss) {
  const sh  = ss.getSheetByName(CONFIG.LEADERBOARD_TAB);
  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return {};

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data    = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, 4).getValues();

  const result = {};
  let lastRank = 1;

  for (const row of data) {
    const rankVal = row[CONFIG.COL_RANK - 1];
    const name    = String(row[CONFIG.COL_NAME - 1]).trim();
    if (!name) continue;

    if (rankVal !== '=' && !isNaN(Number(rankVal)) && rankVal !== '') {
      lastRank = Number(rankVal);
    }
    result[normName(name)] = { name, rank: lastRank };
  }
  return result;
}

function getFirstSubmitter(votes) {
  const withTs = votes.filter(v => v.timestamp);
  if (!withTs.length) return null;
  return withTs.reduce((earliest, v) => new Date(v.timestamp) < new Date(earliest.timestamp) ? v : earliest);
}

function applyScores(ss, votes, correct, firstSubmitter, firstGotItRight, rankSnapshot) {
  const sh      = ss.getSheetByName(CONFIG.LEADERBOARD_TAB);
  const lastRow = sh.getLastRow();
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data    = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, 4).getValues();

  const scores = {};
  for (const row of data) {
    const name  = String(row[CONFIG.COL_NAME - 1]).trim();
    const score = Number(row[CONFIG.COL_SCORE - 1]) || 0;
    if (!name) continue;
    scores[normName(name)] = { name, score };
  }

  const maxRank = Object.values(rankSnapshot).reduce((m, e) => Math.max(m, e.rank), 0);

  for (const vote of votes) {
    const isCorrect = vote.answer.trim().toLowerCase() === correct;
    
    // Name is already canonicalized from checkPolls
    const key = normName(vote.name);

    if (!scores[key]) scores[key] = { name: vote.name, score: 0 };
    if (!isCorrect) continue;

    const snapEntry = rankSnapshot[key];
    const rank      = snapEntry ? snapEntry.rank : maxRank + 1;

    let points = rank;

    // Check if they get the first-to-answer bonus
    if (firstGotItRight && firstSubmitter && normName(firstSubmitter.name) === key) {
      points += rank;
    }
    scores[key].score += points;
  }

  const sorted = Object.values(scores).sort((a, b) => b.score !== a.score ? b.score - a.score : a.name.localeCompare(b.name));

  const rankCol  = []; 
  const nameCol  = []; 
  const scoreCol = []; 

  const uniqueScores = [...new Set(sorted.map(p => p.score))].sort((a, b) => b - a);
  sorted.forEach((player) => {
    const rank = uniqueScores.filter(s => s > player.score).length + 1;
    rankCol.push([rank]);
    nameCol.push([player.name]);
    scoreCol.push([player.score]);
  });

  const totalRows = Math.max(numRows, sorted.length);

  if (sorted.length > 0) {
    sh.getRange(CONFIG.DATA_START_ROW, 1, totalRows, 1).clearContent(); 
    sh.getRange(CONFIG.DATA_START_ROW, 3, totalRows, 2).clearContent(); 

    sh.getRange(CONFIG.DATA_START_ROW, 1, rankCol.length,  1).setValues(rankCol);   
    sh.getRange(CONFIG.DATA_START_ROW, 3, nameCol.length,  1).setValues(nameCol);   
    sh.getRange(CONFIG.DATA_START_ROW, 4, scoreCol.length, 1).setValues(scoreCol);  
  }
}


// ============================================================
// LOGS, BACKUPS, AND UNDO MANAGEMENT
// ============================================================

function writePollLog(ss, weekLabel, pollId, votes, correctAnswer, firstSubmitter, rankSnapshot) {
  const sh = getOrCreateSheet(ss, CONFIG.POLL_LOG_TAB, [
    'Timestamp', 'Week Label', 'Poll ID', 'Voter Name', 'Pre-Poll Rank',
    'Their Answer', 'Correct Answer', 'Result', 'First?', 'Points Awarded', 'Submitted At'
  ]);
  const now     = new Date().toISOString();
  const correct = String(correctAnswer).trim().toLowerCase();
  const maxRank = Object.values(rankSnapshot).reduce((m, e) => Math.max(m, e.rank), 0);

  const rows = votes.map(v => {
    const key       = normName(v.name);
    const isCorrect = v.answer.trim().toLowerCase() === correct;
    const snapEntry = rankSnapshot[key];
    const rank      = snapEntry ? snapEntry.rank : maxRank + 1;
    const isFirst   = firstSubmitter && normName(firstSubmitter.name) === key;
    const gotBonus  = isFirst && isCorrect;

    let points = 0;
    if (isCorrect) points = rank + (gotBonus ? rank : 0);

    return [now, weekLabel, pollId, v.name, rank, v.answer, correctAnswer, isCorrect ? 'Correct' : 'Wrong', isFirst ? 'First' : '', points, v.timestamp || ''];
  });

  if (rows.length > 0) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
}

function saveSnapshot(ss, pollId, weekLabel) {
  const lbSh   = ss.getSheetByName(CONFIG.LEADERBOARD_TAB);
  const snapSh = getOrCreateSheet(ss, CONFIG.SNAPSHOTS_TAB, ['Snapshot At', 'Poll ID', 'Week Label', 'Row', 'Rank', 'Name', 'Score']);

  const lastRow = lbSh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data    = lbSh.getRange(CONFIG.DATA_START_ROW, 1, numRows, 4).getValues();
  const now     = new Date().toISOString();

  const rows = data
    .filter(row => String(row[CONFIG.COL_NAME - 1]).trim())
    .map((row, idx) => [
      now, pollId, weekLabel, CONFIG.DATA_START_ROW + idx, 
      row[CONFIG.COL_RANK - 1], row[CONFIG.COL_NAME - 1], row[CONFIG.COL_SCORE - 1]
    ]);

  if (rows.length > 0) snapSh.getRange(snapSh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
}

function undoLastPoll() {
  const ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const snapSh = ss.getSheetByName(CONFIG.SNAPSHOTS_TAB);
  if (!snapSh || snapSh.getLastRow() < 2) return;

  const data     = snapSh.getRange(2, 1, snapSh.getLastRow() - 1, 7).getValues();
  const lastPoll = data[data.length - 1][1]; 
  const snapRows = data.filter(row => row[1] === lastPoll);
  const lbSh = ss.getSheetByName(CONFIG.LEADERBOARD_TAB);

  for (const row of snapRows) {
    const sheetRow = row[3]; 
    lbSh.getRange(sheetRow, CONFIG.COL_RANK,  1, 1).setValue(row[4]); 
    lbSh.getRange(sheetRow, CONFIG.COL_NAME,  1, 1).setValue(row[5]); 
    lbSh.getRange(sheetRow, CONFIG.COL_SCORE, 1, 1).setValue(row[6]); 
  }

  const configSh = ss.getSheetByName(CONFIG.CONFIG_TAB);
  if (configSh && configSh.getLastRow() > 1) {
    const cfgData = configSh.getRange(2, 1, configSh.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < cfgData.length; i++) {
      if (String(cfgData[i][0]).trim() === String(lastPoll).trim()) {
        configSh.getRange(i + 2, 4).setValue('FALSE');
        configSh.getRange(i + 2, 5).setValue('');
        break;
      }
    }
  }

  for (let i = snapSh.getLastRow(); i >= 2; i--) {
    if (String(snapSh.getRange(i, 2).getValue()).trim() === String(lastPoll).trim()) snapSh.deleteRow(i);
  }
}

function exportLeaderboardPDF(ss, weekLabel) {
  // Allows manual execution from Apps Script editor drop-down or automated checkPolls call
  ss        = ss || SpreadsheetApp.openById(CONFIG.SHEET_ID);
  weekLabel = weekLabel || new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  
  const sh  = ss.getSheetByName(CONFIG.LEADERBOARD_TAB);
  if (!sh) throw new Error('Leaderboard tab not found: ' + CONFIG.LEADERBOARD_TAB);
  
  const gid = sh.getSheetId();

  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/export'
    + '?exportFormat=pdf&format=pdf&gid=' + gid + '&range=' + encodeURIComponent(CONFIG.PDF_RANGE)
    + '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) throw new Error('PDF conversion engine failed with HTTP ' + response.getResponseCode());

  const blob = response.getBlob().setName('Fast Facts Leaderboard — ' + weekLabel + '.pdf');
  const folders = DriveApp.getFoldersByName(CONFIG.PDF_FOLDER_NAME);
  const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.PDF_FOLDER_NAME);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}


// ============================================================
// SYSTEM UTILITIES
// ============================================================

function normName(name) {
  return String(name).toLowerCase().trim();
}

function loadAliases(ss) {
  const sh = ss.getSheetByName(CONFIG.ALIASES_TAB);
  if (!sh || sh.getLastRow() < 2) return {};

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const aliases = {};
  for (const [voterName, canonicalName] of data) {
    if (voterName && canonicalName) aliases[normName(voterName)] = String(canonicalName).trim();
  }
  return aliases;
}

function fuzzyMatchName(voterName, rankSnapshot, aliases) {
  const norm = normName(voterName);
  if (aliases && aliases[norm]) {
    const aliasTarget = normName(aliases[norm]);
    if (rankSnapshot[aliasTarget]) return rankSnapshot[aliasTarget].name;
    return aliases[norm];
  }

  if (rankSnapshot[norm]) return rankSnapshot[norm].name;
  const entries = Object.entries(rankSnapshot);

  for (const [key, val] of entries) {
    if (key.startsWith(norm) || norm.startsWith(key)) return val.name;
  }

  const voterWords = norm.split(/\s+/).filter(Boolean);
  for (const [key, val] of entries) {
    const leaderWords = key.split(/\s+/).filter(Boolean);
    const shorter = voterWords.length <= leaderWords.length ? voterWords : leaderWords;
    const longer  = voterWords.length <= leaderWords.length ? leaderWords : voterWords;
    if (shorter.every(w => longer.some(l => l.startsWith(w) || w.startsWith(l)))) return val.name;
  }
  return null; 
}

function getOrCreateSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function log(sheet, message) {
  if (!sheet) return;
  sheet.appendRow([new Date().toISOString(), message]);
}

function fixRanks() {
  const ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh  = ss.getSheetByName(CONFIG.LEADERBOARD_TAB);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const data    = sh.getRange(CONFIG.DATA_START_ROW, 1, numRows, 4).getValues();

  const players = data.map(row => ({ name: String(row[2]).trim(), score: Number(row[3]) || 0 })).filter(p => p.name);
  const uniqueScores = [...new Set(players.map(p => p.score))].sort((a, b) => b - a);
  const rankCol = players.map(player => [uniqueScores.filter(s => s > player.score).length + 1]);

  sh.getRange(CONFIG.DATA_START_ROW, 1, rankCol.length, 1).setValues(rankCol);
}

function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  getOrCreateSheet(ss, CONFIG.CONFIG_TAB,     ['Poll ID', 'Week Label', 'Correct Answer', 'Processed', 'Processed At']);
  getOrCreateSheet(ss, CONFIG.POLL_LOG_TAB,   ['Timestamp', 'Week Label', 'Poll ID', 'Voter Name', 'Pre-Poll Rank', 'Their Answer', 'Correct Answer', 'Result', 'First?', 'Points Awarded', 'Submitted At']);
  getOrCreateSheet(ss, CONFIG.SCRIPT_LOG_TAB, ['Timestamp', 'Message']);
  getOrCreateSheet(ss, CONFIG.ALIASES_TAB,     ['Voter Name (as typed in StrawPoll)', 'Canonical Name (as on leaderboard)']);
  getOrCreateSheet(ss, CONFIG.SNAPSHOTS_TAB,   ['Snapshot At', 'Poll ID', 'Week Label', 'Row', 'Rank', 'Name', 'Score']);
}

// ============================================================
// UTILITY: FULL LEADERBOARD RECALCULATION (WITH HISTORICAL BASELINE)
// ============================================================

function recalculateAllPolls() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  
  // 1. Create hidden backups of the current state BEFORE wiping
  backupState(); 
  
  const configSh = ss.getSheetByName(CONFIG.CONFIG_TAB);
  const snapSh   = ss.getSheetByName(CONFIG.SNAPSHOTS_TAB);
  const lbSh     = ss.getSheetByName(CONFIG.LEADERBOARD_TAB);
  const logSh    = ss.getSheetByName(CONFIG.POLL_LOG_TAB);
  
  // 2. Find the historical baseline from the very first automated poll
  let baselineData = [];
  if (configSh && snapSh && configSh.getLastRow() > 1 && snapSh.getLastRow() > 1) {
    // Get the first Poll ID from the top of the Config list
    const firstPollId = String(configSh.getRange(2, 1).getValue()).trim();
    
    // Grab the snapshot taken right before that first poll was scored
    const snapData = snapSh.getDataRange().getValues();
    baselineData = snapData.filter(row => String(row[1]).trim() === firstPollId);
  }
  
  if (baselineData.length === 0) {
    throw new Error("Could not find the snapshot for the first poll. Run restoreState() and check your Config tab.");
  }
  
  // 3. Wipe the current leaderboard
  const lastRowLb = lbSh.getLastRow();
  if (lastRowLb >= CONFIG.DATA_START_ROW) {
    lbSh.getRange(CONFIG.DATA_START_ROW, 1, lastRowLb - CONFIG.DATA_START_ROW + 1, 4).clearContent();
  }
  
  // 4. Re-populate the leaderboard with the historical baseline
  for (const row of baselineData) {
    const sheetRow = row[3]; // The row number from the snapshot
    lbSh.getRange(sheetRow, CONFIG.COL_RANK, 1, 1).setValue(row[4]); 
    lbSh.getRange(sheetRow, CONFIG.COL_NAME, 1, 1).setValue(row[5]); 
    lbSh.getRange(sheetRow, CONFIG.COL_SCORE, 1, 1).setValue(row[6]); 
  }
  
  // 5. Wipe the Poll Log (leaving headers intact)
  if (logSh && logSh.getLastRow() > 1) {
    logSh.getRange(2, 1, logSh.getLastRow() - 1, logSh.getLastColumn()).clearContent();
  }

  // 6. Wipe the Snapshots tab (leaving headers intact)
  if (snapSh && snapSh.getLastRow() > 1) {
    snapSh.getRange(2, 1, snapSh.getLastRow() - 1, snapSh.getLastColumn()).clearContent();
  }

  // 7. Reset the Config tab flags so checkPolls thinks they are all new
  const lastRowCfg = configSh.getLastRow();
  if (lastRowCfg > 1) {
    configSh.getRange(2, 4, lastRowCfg - 1, 1).setValue('FALSE');
    configSh.getRange(2, 5, lastRowCfg - 1, 1).setValue('');
  }

  // Force Google Sheets to commit the baseline and wipe before processing
  SpreadsheetApp.flush();

  // 8. Trigger the main engine to process every poll sequentially from the restored baseline
  checkPolls();
}

// ============================================================
// SAFETY: BACKUP AND RESTORE
// ============================================================

function backupState() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const tabsToBackup = [CONFIG.LEADERBOARD_TAB, CONFIG.POLL_LOG_TAB, CONFIG.SNAPSHOTS_TAB, CONFIG.CONFIG_TAB];
  
  // 1. Delete any old backups to prevent clutter
  for (const sheet of ss.getSheets()) {
    if (sheet.getName().endsWith('_BACKUP')) {
      ss.deleteSheet(sheet);
    }
  }
  
  // 2. Duplicate current tabs, rename them, and hide them
  for (const tabName of tabsToBackup) {
    const sh = ss.getSheetByName(tabName);
    if (sh) {
      const backup = sh.copyTo(ss);
      backup.setName(tabName + '_BACKUP');
      backup.hideSheet(); 
    }
  }
  SpreadsheetApp.flush();
}

function restoreState() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const tabsToRestore = [CONFIG.LEADERBOARD_TAB, CONFIG.POLL_LOG_TAB, CONFIG.SNAPSHOTS_TAB, CONFIG.CONFIG_TAB];
  
  let restoredCount = 0;
  
  for (const tabName of tabsToRestore) {
    const backupSh = ss.getSheetByName(tabName + '_BACKUP');
    if (backupSh) {
      // Delete the active (newly calculated) sheet
      const activeSh = ss.getSheetByName(tabName);
      if (activeSh) ss.deleteSheet(activeSh);
      
      // Restore the backup to active duty
      backupSh.setName(tabName);
      backupSh.showSheet();
      restoredCount++;
    }
  }
  
  if (restoredCount > 0) {
    SpreadsheetApp.getUi().alert('Success: Leaderboard and logs restored to their previous state.');
  } else {
    SpreadsheetApp.getUi().alert('Error: No backup found. Did you run the backup first?');
  }
}
