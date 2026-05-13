// ============================================================
// IGI School of Gemology — RSP M1 Assessment
// Google Apps Script Backend
// Deploy as Web App: Execute as Me, Anyone can access
// ============================================================

const SHEET_ID = "1GN5J4R7jePjShdsGxuokBRwwnaR7QlCVoLMh__9hUlI";
const TRAINER_USERS = { "admin": "IGIRSP2026", "john": "IGIJohn2026" };
const SESSION_SECRET = "IGI_RSP_SESSION_2026";

function getSpreadsheet() { return SpreadsheetApp.openById(SHEET_ID); }

// ── Sheet bootstrap ──────────────────────────────────────────
function ensureSheets() {
  const ss = getSpreadsheet();
  const needed = ["Config","Responses","By_Brand","By_Branch","By_Batch","Individual_Progress"];
  needed.forEach(name => { if (!ss.getSheetByName(name)) ss.insertSheet(name); });

  // Config headers
  const cfg = ss.getSheetByName("Config");
  if (cfg.getLastRow() === 0) {
    cfg.appendRow(["BatchCode","IGICentre","JewelryBrand","AssessmentType","Active","CreatedAt","CreatedBy","ParticipantCount"]);
  }

  // Responses headers
  const res = ss.getSheetByName("Responses");
  if (res.getLastRow() === 0) {
    const headers = ["Timestamp","BatchCode","IGICentre","JewelryBrand","StoreBranch",
      "Name","Mobile","Email","Designation","Experience",
      "Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8","Q9","Q10",
      "Q11","Q12","Q13","Q14","Q15","Q16","Q17","Q18","Q19","Q20",
      "Q21","Q22","Q23","Q24","Q25",
      "ND","TE","EI","CS","RO","PV","CA","Overall","Persona",
      "ND_Flag","TE_Flag","EI_Flag","CS_Flag","RO_Flag","PV_Flag","Consistency_Overall",
      "Commitment45Day"
    ];
    res.appendRow(headers);
    res.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#094d59").setFontColor("#ffffff");
  }
}

// ── CORS helper ──────────────────────────────────────────────
function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Session token ────────────────────────────────────────────
function makeToken(username) {
  return Utilities.base64Encode(username + "|" + SESSION_SECRET + "|" + new Date().toDateString());
}
function validateToken(token) {
  if (!token) return false;
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split("|");
    return parts[1] === SESSION_SECRET && parts[2] === new Date().toDateString();
  } catch(e) { return false; }
}

// ── doGet router ─────────────────────────────────────────────
function doGet(e) {
  ensureSheets();
  const action = e.parameter.action || "";

  try {
    if (action === "login")           return handleLogin(e);
    if (action === "getConfig")       return handleGetConfig(e);
    if (action === "checkMobile")     return handleCheckMobile(e);
    if (action === "activateBatch")   return handleActivateBatch(e);
    if (action === "deactivateBatch") return handleDeactivateBatch(e);
    if (action === "getBatchStatus")  return handleGetBatchStatus(e);
    if (action === "getResults")      return handleGetResults(e);
    if (action === "getAllBatches")    return handleGetAllBatches(e);
    return corsOutput({ ok: false, error: "Unknown action" });
  } catch(err) {
    return corsOutput({ ok: false, error: err.toString() });
  }
}

// ── doPost — save assessment result ─────────────────────────
function doPost(e) {
  ensureSheets();
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSpreadsheet();
    const res = ss.getSheetByName("Responses");

    const row = [
      new Date().toISOString(),
      data.batchCode, data.igiCentre, data.jewelryBrand, data.storeBranch,
      data.name, data.mobile, data.email, data.designation, data.experience,
      data.q1,data.q2,data.q3,data.q4,data.q5,
      data.q6,data.q7,data.q8,data.q9,data.q10,
      data.q11,data.q12,data.q13,data.q14,data.q15,
      data.q16,data.q17,data.q18,data.q19,data.q20,
      data.q21,data.q22,data.q23,data.q24,data.q25,
      data.nd, data.te, data.ei, data.cs, data.ro, data.pv, data.ca,
      data.overall, data.persona,
      data.nd_flag, data.te_flag, data.ei_flag,
      data.cs_flag, data.ro_flag, data.pv_flag, data.consistency_overall,
      data.commitment
    ];
    res.appendRow(row);

    // Update participant count in Config
    const cfg = ss.getSheetByName("Config");
    const cfgData = cfg.getDataRange().getValues();
    for (let i = 1; i < cfgData.length; i++) {
      if (cfgData[i][0] === data.batchCode && cfgData[i][4] === "Y") {
        const currentCount = cfgData[i][7] || 0;
        cfg.getRange(i+1, 8).setValue(Number(currentCount)+1);
        break;
      }
    }
    return corsOutput({ ok: true });
  } catch(err) {
    return corsOutput({ ok: false, error: err.toString() });
  }
}

// ── Login ────────────────────────────────────────────────────
function handleLogin(e) {
  const u = e.parameter.username || "";
  const p = e.parameter.password || "";
  if (TRAINER_USERS[u] && TRAINER_USERS[u] === p) {
    return corsOutput({ ok: true, token: makeToken(u), username: u });
  }
  return corsOutput({ ok: false, error: "Invalid credentials" });
}

// ── Get active batch config (for participant tool) ───────────
function handleGetConfig(e) {
  const cfg = getSpreadsheet().getSheetByName("Config");
  const rows = cfg.getDataRange().getValues();
  for (let i = rows.length-1; i >= 1; i--) {
    if (rows[i][4] === "Y") {
      return corsOutput({
        ok: true,
        batchCode: rows[i][0], igiCentre: rows[i][1],
        jewelryBrand: rows[i][2], assessmentType: rows[i][3]
      });
    }
  }
  return corsOutput({ ok: false, error: "No active batch" });
}

// ── Check mobile (one attempt rule) ─────────────────────────
function handleCheckMobile(e) {
  const mobile = e.parameter.mobile || "";
  const batchCode = e.parameter.batchCode || "";
  const res = getSpreadsheet().getSheetByName("Responses");
  const rows = res.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][6]) === String(mobile) && rows[i][1] === batchCode) {
      return corsOutput({ ok: true, attempted: true, name: rows[i][5] });
    }
  }
  return corsOutput({ ok: true, attempted: false });
}

// ── Activate batch (trainer) ─────────────────────────────────
function handleActivateBatch(e) {
  if (!validateToken(e.parameter.token)) return corsOutput({ ok:false, error:"Unauthorized" });
  const cfg = getSpreadsheet().getSheetByName("Config");
  const rows = cfg.getDataRange().getValues();
  // Deactivate all current
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][4] === "Y") cfg.getRange(i+1,5).setValue("N");
  }
  cfg.appendRow([
    e.parameter.batchCode, e.parameter.igiCentre, e.parameter.jewelryBrand,
    e.parameter.assessmentType, "Y", new Date().toISOString(),
    e.parameter.username || "admin", 0
  ]);
  return corsOutput({ ok: true });
}

// ── Deactivate batch (trainer) ───────────────────────────────
function handleDeactivateBatch(e) {
  if (!validateToken(e.parameter.token)) return corsOutput({ ok:false, error:"Unauthorized" });
  const cfg = getSpreadsheet().getSheetByName("Config");
  const rows = cfg.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === e.parameter.batchCode) {
      cfg.getRange(i+1,5).setValue("N");
      return corsOutput({ ok: true });
    }
  }
  return corsOutput({ ok: false, error: "Batch not found" });
}

// ── Get batch live status ────────────────────────────────────
function handleGetBatchStatus(e) {
  if (!validateToken(e.parameter.token)) return corsOutput({ ok:false, error:"Unauthorized" });
  const batchCode = e.parameter.batchCode || "";
  const res = getSpreadsheet().getSheetByName("Responses");
  const rows = res.getDataRange().getValues();
  let participants = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === batchCode) {
      participants.push({
        name: rows[i][5], mobile: rows[i][6],
        persona: rows[i][38], overall: rows[i][37],
        timestamp: rows[i][0]
      });
    }
  }
  return corsOutput({ ok: true, count: participants.length, participants: participants });
}

// ── Get full results for dashboard ──────────────────────────
function handleGetResults(e) {
  if (!validateToken(e.parameter.token)) return corsOutput({ ok:false, error:"Unauthorized" });
  const batchCode = e.parameter.batchCode || "";
  const res = getSpreadsheet().getSheetByName("Responses");
  const rows = res.getDataRange().getValues();
  const headers = rows[0];
  let results = [];
  for (let i = 1; i < rows.length; i++) {
    if (!batchCode || rows[i][1] === batchCode) {
      let obj = {};
      headers.forEach((h,idx) => { obj[h] = rows[i][idx]; });
      results.push(obj);
    }
  }
  return corsOutput({ ok: true, results: results });
}

// ── Get all batches ──────────────────────────────────────────
function handleGetAllBatches(e) {
  if (!validateToken(e.parameter.token)) return corsOutput({ ok:false, error:"Unauthorized" });
  const cfg = getSpreadsheet().getSheetByName("Config");
  const rows = cfg.getDataRange().getValues();
  let batches = [];
  for (let i = 1; i < rows.length; i++) {
    batches.push({
      batchCode: rows[i][0], igiCentre: rows[i][1],
      jewelryBrand: rows[i][2], assessmentType: rows[i][3],
      active: rows[i][4], createdAt: rows[i][5],
      participantCount: rows[i][7]
    });
  }
  return corsOutput({ ok: true, batches: batches.reverse() });
}
