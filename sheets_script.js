// ============================================================
// Google Apps Script — Full exam backend
// Handles: questions loading, OTP verification, result saving, video upload
// Deploy → New deployment → Web app → Execute as: Me → Anyone → Deploy
// ============================================================

const OTP_EXPIRY_MINUTES = 10;
const VIDEO_FOLDER_NAME = "ExamVideos";

// ===== GET: handles questions + OTP verification =====
function doGet(e) {
  var action = (e.parameter.action || "").toLowerCase();
  var callback = e.parameter.callback || "handleSheetData";

  try {
    if (action === "sendotp") {
      return jsonp(callback, sendOtp(e.parameter.email, e.parameter.fingerprint));
    }
    if (action === "verifyotp") {
      return jsonp(callback, verifyOtp(e.parameter.email, e.parameter.otp));
    }
    if (action === "checkemail") {
      return jsonp(callback, { used: isEmailUsed(e.parameter.email) });
    }
    // Default: return questions
    return jsonp(callback, getQuestions());
  } catch (err) {
    return jsonp(callback, { error: err.toString() });
  }
}

// ===== Helper: JSONP response =====
function jsonp(callback, data) {
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(data) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ===== Read questions from first sheet =====
function getQuestions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var bank = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var setId = String(row[0] || "").trim();
    var qtype = String(row[1] || "").trim().toLowerCase();
    var question = String(row[2] || "").trim();
    if (!setId || !qtype || !question) continue;

    if (!bank[setId]) bank[setId] = { mcq: [], scenarios: [], project: null };

    if (qtype === "mcq") {
      bank[setId].mcq.push({
        question: question,
        options: [String(row[3]||""), String(row[4]||""), String(row[5]||""), String(row[6]||"")],
        answer: String(row[7] || "A").trim()
      });
    } else if (qtype === "scenario") {
      bank[setId].scenarios.push({
        label: String(row[9] || "").trim() || ("Scenario " + (bank[setId].scenarios.length + 1)),
        context: String(row[8] || "").trim(),
        question: question
      });
    } else if (qtype === "project") {
      bank[setId].project = {
        title: String(row[10] || "").trim() || "Project",
        description: question
      };
    }
  }
  return bank;
}

// ===== Check if email already submitted (and NOT reset) =====
// Looks at the "Reset" column (col 22) — if admin typed YES, allow retake
function isEmailUsed(email) {
  if (!email) return false;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results");
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  var emailLower = email.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][2] || "").toLowerCase().trim();
    var resetFlag = String(data[i][21] || "").toUpperCase().trim(); // column V = index 21
    if (rowEmail === emailLower && resetFlag !== "YES") return true;
  }
  return false;
}

// ===== Send OTP to email =====
function sendOtp(email, fingerprint) {
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, error: "Invalid email" };
  }

  if (isEmailUsed(email)) {
    return { success: false, error: "already_submitted", message: "This email has already taken the exam. Contact the administrator." };
  }

  var otp = Math.floor(100000 + Math.random() * 900000).toString();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("OTPs");
  if (!sheet) {
    sheet = ss.insertSheet("OTPs");
    sheet.appendRow(["Email", "OTP", "Generated At", "Expires At", "Fingerprint", "Verified"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }

  var now = new Date();
  var expiresAt = new Date(now.getTime() + OTP_EXPIRY_MINUTES * 60 * 1000);
  sheet.appendRow([email.toLowerCase().trim(), otp, now, expiresAt, fingerprint || "", "NO"]);

  try {
    MailApp.sendEmail({
      to: email,
      subject: "Your Exam Verification Code",
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:480px;margin:20px auto;padding:24px;background:#f5f7fa;border-radius:12px;">' +
        '<h2 style="color:#3b7fff;margin:0 0 16px;">Exam Access Code</h2>' +
        '<p style="color:#333;font-size:14px;">Use this 6-digit code to start your exam:</p>' +
        '<div style="background:white;padding:20px;text-align:center;border-radius:8px;border:2px solid #3b7fff;margin:16px 0;">' +
        '<span style="font-size:32px;letter-spacing:8px;font-weight:700;color:#111;">' + otp + '</span>' +
        '</div>' +
        '<p style="color:#666;font-size:12px;">This code expires in ' + OTP_EXPIRY_MINUTES + ' minutes. Do not share it with anyone.</p>' +
        '</div>'
    });
    return { success: true, message: "OTP sent to " + email };
  } catch (err) {
    return { success: false, error: "mail_failed", message: err.toString() };
  }
}

// ===== Verify OTP =====
function verifyOtp(email, otp) {
  if (!email || !otp) return { success: false, error: "Missing email or OTP" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("OTPs");
  if (!sheet) return { success: false, error: "No OTP found. Please request a new code." };

  var data = sheet.getDataRange().getValues();
  var emailLower = email.toLowerCase().trim();
  var now = new Date();

  // Find most recent unverified OTP for this email
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    if (String(row[0] || "").toLowerCase().trim() === emailLower && String(row[5]).toUpperCase() === "NO") {
      if (String(row[1]) !== String(otp)) {
        return { success: false, error: "Invalid OTP" };
      }
      var expiresAt = new Date(row[3]);
      if (now > expiresAt) {
        return { success: false, error: "OTP expired. Please request a new code." };
      }
      // Mark as verified
      sheet.getRange(i + 1, 6).setValue("YES");
      return { success: true, message: "OTP verified" };
    }
  }
  return { success: false, error: "No active OTP found. Please request a new code." };
}

// ===== POST: save results OR upload video =====
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Video upload
    if (data.type === "video") {
      return handleVideoUpload(data);
    }

    // Exam results
    return handleResultSubmission(data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===== Save exam results =====
function handleResultSubmission(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results");

  if (!sheet) {
    sheet = ss.insertSheet("Results");
    sheet.appendRow([
      "Timestamp", "Name", "Email", "Set",
      "Score", "Total MCQ", "Percentage", "Result",
      "Correct", "Wrong", "Skipped",
      "Violations", "Time Taken",
      "Scenario 1", "Scenario 2", "Project",
      "AI Violations", "Paste Attempts", "Device Fingerprint", "Video URL", "Question Times",
      "Reset (YES to allow retake)"
    ]);
    sheet.getRange(1, 1, 1, 22).setFontWeight("bold");
    sheet.getRange(1, 22).setBackground("#fff3cd");
  }

  // Auto-add Reset column to existing sheets that don't have it
  if (sheet.getLastColumn() < 22) {
    sheet.getRange(1, 22).setValue("Reset (YES to allow retake)").setFontWeight("bold").setBackground("#fff3cd");
  }

  // Enable text wrap on the full sheet (safe to run every time)
  sheet.getRange(1, 1, sheet.getMaxRows(), 22).setWrap(true);

  var totalMCQ = (data.correct || 0) + (data.wrong || 0) + (data.skipped || 0);
  var result = (data.percentage >= 60) ? "PASSED" : "FAILED";

  sheet.appendRow([
    new Date().toLocaleString(),
    data.name || "",
    data.email || "",
    data.set || "",
    data.score || 0,
    totalMCQ,
    (data.percentage || 0) + "%",
    result,
    data.correct || 0,
    data.wrong || 0,
    data.skipped || 0,
    data.violations || 0,
    data.timeTaken || "",
    data.scenario1 || "",
    data.scenario2 || "",
    data.project || "",
    JSON.stringify(data.aiViolations || []),
    data.pasteAttempts || 0,
    data.deviceFingerprint || "",
    data.videoUrl || "",
    JSON.stringify(data.questionTimes || []),
    "" // Reset column — admin types YES here to allow retake
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== Upload video to Google Drive =====
function handleVideoUpload(data) {
  var base64 = data.video; // data URL without prefix
  var email = data.email || "unknown";
  var name = data.name || "candidate";

  if (!base64) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "No video data" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Find or create ExamVideos folder
  var folders = DriveApp.getFoldersByName(VIDEO_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(VIDEO_FOLDER_NAME);

  // Decode and save
  var decoded = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(decoded, "video/webm", sanitize(name) + "_" + sanitize(email) + "_" + new Date().getTime() + ".webm");
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "success", videoUrl: file.getUrl() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sanitize(str) {
  return String(str || "").replace(/[^a-zA-Z0-9._-]/g, "_").substring(0, 50);
}

// Run this ONCE to authorize Drive access (then delete)
function authorizeDrive() {
  var folders = DriveApp.getFoldersByName("ExamVideos");
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("ExamVideos");
  var testFile = folder.createFile("auth_test.txt", "Drive authorization successful");
  Logger.log("Test file created: " + testFile.getUrl());
}

// Run this ONCE to add the "Reset" column to an existing Results sheet
function addResetColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results");
  if (!sheet) {
    Logger.log("No Results sheet found yet.");
    return;
  }
  sheet.getRange(1, 22).setValue("Reset (YES to allow retake)").setFontWeight("bold").setBackground("#fff3cd");
  Logger.log("Reset column added at column V.");
}

// Run this to enable text wrapping on all Results cells
function enableTextWrap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Results");
  if (!sheet) { Logger.log("No Results sheet"); return; }
  sheet.getRange(1, 1, sheet.getMaxRows(), 22).setWrap(true);
  Logger.log("Text wrap enabled on Results sheet.");
}
