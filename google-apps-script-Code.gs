const SHEET_NAME = "Sheet1";
const REWARDS_SHEET_NAME = "Rewards";
const LEARNING_REWARD = 5;
const MINUTE_REWARD = 10;
const GAME_COST = 20;
const HEADERS = [
  "code",
  "activityName",
  "minutes",
  "timestampISO",
  "timestampDisplay",
  "clientId",
  "userAgent",
  "entryId",
  "topic",
  "skill",
  "type",
  "requestId"
];

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const action = String(p.action || "").toLowerCase();
  const callback = String(p.callback || "").trim();

  try {
    if (action === "ping") {
      return respond({ ok: true, msg: "pong-v6" }, callback);
    }

    const code = sanitizeCode(p.code);
    if (!code) return respond({ ok: false, error: "Missing code" }, callback);

    const sheet = getSheet();
    ensureHeaders(sheet);

    if (action === "load") {
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        return respond({ ok: true, rows: loadRows(sheet, code), rewards: getRewards(code, sheet) }, callback);
      } finally {
        lock.releaseLock();
      }
    }

    if (action === "save") {
      const topic = cleanText(p.topic, 200);
      const skill = cleanText(p.skill, 200);
      const type = cleanText(p.type, 200);
      const requestId = cleanText(p.requestId, 100);
      const minutes = Number(p.minutes || 0);

      if (!topic || !skill || !type) {
        return respond({ ok: false, error: "Missing topic, skill, or type" }, callback);
      }
      if (!Number.isInteger(minutes) || minutes < 1 || minutes > 1440) {
        return respond({ ok: false, error: "Minutes must be a whole number from 1 to 1440" }, callback);
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const existingRow = requestId ? findRowByRequestId(sheet, code, requestId) : null;
        if (existingRow) {
          return respond({
            ok: true,
            row: existingRow,
            rewards: getRewards(code, sheet),
            duplicate: true
          }, callback);
        }

        const rewardRecord = getOrCreateRewardRecord(code, sheet);
        const now = new Date();
        const entryId = Utilities.getUuid();
        const timestampISO = now.toISOString();
        const timestampDisplay = Utilities.formatDate(
          now,
          Session.getScriptTimeZone(),
          "MMM d, yyyy h:mm a"
        );

        sheet.appendRow([
          code,
          topic,
          minutes || "",
          timestampISO,
          timestampDisplay,
          cleanText(p.clientId, 200),
          cleanText(p.userAgent, 300),
          entryId,
          topic,
          skill,
          type,
          requestId
        ]);

        const rewards = recordCompletedActivity(rewardRecord, minutes);

        return respond({
          ok: true,
          row: { entryId, topic, skill, type, minutes: minutes || "", timestampISO, timestampDisplay },
          rewards: rewards
        }, callback);
      } finally {
        lock.releaseLock();
      }
    }

    if (action === "delete") {
      const entryId = cleanText(p.entryId, 100);
      if (!entryId) return respond({ ok: false, error: "Missing entryId" }, callback);

      const deleted = deleteEntry(sheet, code, entryId);
      return respond({ ok: deleted, error: deleted ? "" : "Entry not found" }, callback);
    }

    if (action === "reset") {
      deleteRowsForCode(sheet, code);
      return respond({ ok: true }, callback);
    }

    if (action === "rewards") {
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        return respond({ ok: true, rewards: getRewards(code, sheet) }, callback);
      } finally {
        lock.releaseLock();
      }
    }

    if (action === "play") {
      const game = cleanText(p.game, 50);
      if (!game) return respond({ ok: false, error: "Missing game" }, callback);

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const rewards = spendPointsForGame(code, game, sheet);
        return respond({ ok: true, rewards: rewards }, callback);
      } finally {
        lock.releaseLock();
      }
    }

    return respond({ ok: false, error: "Invalid action" }, callback);
  } catch (err) {
    return respond({ ok: false, error: String(err) }, callback);
  }
}

function getRewardsSheet() {
  const book = SpreadsheetApp.getActive();
  let sheet = book.getSheetByName(REWARDS_SHEET_NAME);
  if (!sheet) sheet = book.insertSheet(REWARDS_SHEET_NAME);

  const headers = [
    "code", "points", "lifetimeActivities", "lifetimeMinutes",
    "learningMilestones", "minuteMilestones", "gamesPlayed", "updatedISO"
  ];
  headers.forEach(function(header, index) {
    if (!sheet.getRange(1, index + 1).getValue()) sheet.getRange(1, index + 1).setValue(header);
  });
  return sheet;
}

function findRewardRow(rewardsSheet, code) {
  const lastRow = rewardsSheet.getLastRow();
  if (lastRow < 2) return 0;
  const codes = rewardsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < codes.length; i++) {
    if (String(codes[i][0] || "") === code) return i + 2;
  }
  return 0;
}

function initialLifetimeStats(activitySheet, code) {
  const rows = loadRows(activitySheet, code);
  return {
    activities: rows.length,
    minutes: rows.reduce(function(sum, row) { return sum + (Number(row.minutes) || 0); }, 0)
  };
}

function getOrCreateRewardRecord(code, activitySheet) {
  const rewardsSheet = getRewardsSheet();
  let rowNumber = findRewardRow(rewardsSheet, code);

  if (!rowNumber) {
    const stats = initialLifetimeStats(activitySheet, code);
    const learningMilestones = Math.floor(stats.activities / 5);
    const minuteMilestones = Math.floor(stats.minutes / 200);
    const points = learningMilestones * LEARNING_REWARD + minuteMilestones * MINUTE_REWARD;
    rewardsSheet.appendRow([
      code, points, stats.activities, stats.minutes,
      learningMilestones, minuteMilestones, 0, new Date().toISOString()
    ]);
    rowNumber = rewardsSheet.getLastRow();
  }

  const row = rewardsSheet.getRange(rowNumber, 1, 1, 8).getValues()[0];
  return {
    sheet: rewardsSheet,
    rowNumber: rowNumber,
    code: code,
    points: Number(row[1]) || 0,
    lifetimeActivities: Number(row[2]) || 0,
    lifetimeMinutes: Number(row[3]) || 0,
    learningMilestones: Number(row[4]) || 0,
    minuteMilestones: Number(row[5]) || 0,
    gamesPlayed: Number(row[6]) || 0
  };
}

function rewardSummary(record) {
  return {
    points: record.points,
    lifetimeActivities: record.lifetimeActivities,
    lifetimeMinutes: record.lifetimeMinutes,
    learningMilestones: record.learningMilestones,
    minuteMilestones: record.minuteMilestones,
    gamesPlayed: record.gamesPlayed,
    learningReward: LEARNING_REWARD,
    minuteReward: MINUTE_REWARD,
    gameCost: GAME_COST
  };
}

function saveRewardRecord(record) {
  record.sheet.getRange(record.rowNumber, 1, 1, 8).setValues([[
    record.code,
    record.points,
    record.lifetimeActivities,
    record.lifetimeMinutes,
    record.learningMilestones,
    record.minuteMilestones,
    record.gamesPlayed,
    new Date().toISOString()
  ]]);
}

function getRewards(code, activitySheet) {
  return rewardSummary(getOrCreateRewardRecord(code, activitySheet));
}

function recordCompletedActivity(record, minutes) {
  record.lifetimeActivities += 1;
  record.lifetimeMinutes += Number(minutes) || 0;

  const learningMilestones = Math.floor(record.lifetimeActivities / 5);
  const minuteMilestones = Math.floor(record.lifetimeMinutes / 200);
  const newLearning = Math.max(0, learningMilestones - record.learningMilestones);
  const newMinute = Math.max(0, minuteMilestones - record.minuteMilestones);

  record.points += newLearning * LEARNING_REWARD + newMinute * MINUTE_REWARD;
  record.learningMilestones = learningMilestones;
  record.minuteMilestones = minuteMilestones;
  saveRewardRecord(record);
  return rewardSummary(record);
}

function spendPointsForGame(code, game, activitySheet) {
  const record = getOrCreateRewardRecord(code, activitySheet);
  if (record.points < GAME_COST) throw new Error("You need 20 points to play this game");
  record.points -= GAME_COST;
  record.gamesPlayed += 1;
  saveRewardRecord(record);
  return rewardSummary(record);
}

function getSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet tab "' + SHEET_NAME + '" not found');
  return sheet;
}

function ensureHeaders(sheet) {
  const width = Math.max(sheet.getLastColumn(), HEADERS.length);
  const current = sheet.getRange(1, 1, 1, width).getValues()[0];
  HEADERS.forEach(function(header, index) {
    if (!current[index]) sheet.getRange(1, index + 1).setValue(header);
  });
}

function loadRows(sheet, code) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues()
    .filter(function(row) { return String(row[0] || "") === code; })
    .map(function(row) {
      const legacyName = String(row[1] || "");
      return {
        entryId: String(row[7] || "legacy-" + code + "-" + String(row[3] || "")),
        topic: String(row[8] || legacyName || "Legacy activity"),
        skill: String(row[9] || "Unspecified"),
        type: String(row[10] || "Unspecified"),
        minutes: row[2] === "" ? "" : Number(row[2]) || "",
        timestampISO: row[3] instanceof Date ? row[3].toISOString() : String(row[3] || ""),
        timestampDisplay: String(row[4] || "")
      };
    });
}

function findRowByRequestId(sheet, code, requestId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (String(row[0] || "") === code && String(row[11] || "") === requestId) {
      const legacyName = String(row[1] || "");
      return {
        entryId: String(row[7] || "legacy-" + code + "-" + String(row[3] || "")),
        topic: String(row[8] || legacyName || "Legacy activity"),
        skill: String(row[9] || "Unspecified"),
        type: String(row[10] || "Unspecified"),
        minutes: row[2] === "" ? "" : Number(row[2]) || "",
        timestampISO: row[3] instanceof Date ? row[3].toISOString() : String(row[3] || ""),
        timestampDisplay: String(row[4] || "")
      };
    }
  }
  return null;
}

function deleteEntry(sheet, code, entryId) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const rowCode = String(data[i][0] || "");
    const storedId = String(data[i][7] || "");
    const legacyId = "legacy-" + code + "-" + String(data[i][3] || "");
    if (rowCode === code && (storedId === entryId || legacyId === entryId)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function deleteRowsForCode(sheet, code) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0] || "") === code) sheet.deleteRow(i + 1);
  }
}

function sanitizeCode(code) {
  return String(code || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9\-]/g, "")
    .slice(0, 32);
}

function cleanText(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

function respond(obj, callbackName) {
  const jsonText = JSON.stringify(obj);
  if (callbackName) {
    const safeCallback = callbackName.replace(/[^a-zA-Z0-9_$\.]/g, "");
    return ContentService.createTextOutput(safeCallback + "(" + jsonText + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonText)
    .setMimeType(ContentService.MimeType.JSON);
}
