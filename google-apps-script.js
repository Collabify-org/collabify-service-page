/**
 * ============================================================
 *  CollabifySpace — Form Submission Handler
 *  Google Apps Script (paste into script.google.com)
 * ============================================================
 *
 *  WHAT THIS DOES:
 *  1. Receives form data via HTTP POST from your website
 *  2. Appends a new row to your Google Sheet
 *  3. Emails YOU with full submission details
 *  4. Sends a confirmation email to the user
 *
 *  SETUP INSTRUCTIONS: see SETUP-GUIDE.md
 * ============================================================
 */

// ── CONFIGURATION — Edit these values ────────────────────────
var CONFIG = {
  OWNER_EMAIL:    "your@email.com",          // ← Your email address
  SHEET_NAME:     "Collabify Submissions",    // ← Sheet tab name (auto-created)
  BUSINESS_NAME:  "CollabifySpace",
  WEBSITE_URL:    "https://collabifyspace.com"
};
// ─────────────────────────────────────────────────────────────


/**
 * Handles GET requests (used to verify the deployment is live).
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "CollabifySpace endpoint is live." }))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Handles POST requests sent from the website contact form.
 */
function doPost(e) {
  try {
    // 1. Parse incoming JSON body
    var data = JSON.parse(e.postData.contents);

    // 2. Write to Google Sheet
    appendToSheet(data);

    // 3. Email the owner
    sendOwnerEmail(data);

    // 4. Email the user (confirmation)
    if (data.email) {
      sendUserConfirmationEmail(data);
    }

    // 5. Return success to the browser
    return buildResponse({ success: true });

  } catch (err) {
    Logger.log("Error: " + err.toString());
    return buildResponse({ success: false, error: err.toString() });
  }
}


// ── SHEET HELPER ──────────────────────────────────────────────

function appendToSheet(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // Create the sheet if it doesn't exist yet
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);

    // Add header row
    var headers = [
      "Timestamp",
      "Form Type",
      "Full Name",
      "Email",
      "Phone",
      "Company",
      "Service Interested In",
      "Project Stage",
      "Budget Range",
      "Expected Timeline",
      "Project Description",
      "Lead Source",
      "Page Source"
    ];
    sheet.appendRow(headers);

    // Style header row
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1a6fe8");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);

    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
  }

  // Append the data row
  sheet.appendRow([
    data.timestamp    || new Date().toISOString(),
    data.form_type    || "Project Inquiry",
    data.full_name    || "",
    data.email        || "",
    data.phone        || "",
    data.company      || "",
    data.service      || "",
    data.project_stage || "",
    data.budget       || "",
    data.timeline     || "",
    data.description  || "",
    data.lead_source  || "Website",
    data.page_source  || "Contact Page"
  ]);
}


// ── OWNER EMAIL ───────────────────────────────────────────────

function sendOwnerEmail(data) {
  var subject = "🔔 New Inquiry — " + (data.full_name || "Unknown") +
                " | " + (data.service || "General") +
                " | " + CONFIG.BUSINESS_NAME;

  var body =
    "You have received a new project inquiry from your website.\n\n" +
    "══════════════════════════════════════\n" +
    "  CONTACT DETAILS\n" +
    "══════════════════════════════════════\n" +
    "Name:        " + (data.full_name || "—") + "\n" +
    "Email:       " + (data.email     || "—") + "\n" +
    "Phone:       " + (data.phone     || "—") + "\n" +
    "Company:     " + (data.company   || "—") + "\n\n" +
    "══════════════════════════════════════\n" +
    "  PROJECT DETAILS\n" +
    "══════════════════════════════════════\n" +
    "Service:     " + (data.service       || "—") + "\n" +
    "Stage:       " + (data.project_stage  || "—") + "\n" +
    "Budget:      " + (data.budget        || "—") + "\n" +
    "Timeline:    " + (data.timeline      || "—") + "\n\n" +
    "──────────────────────────────────────\n" +
    "PROJECT DESCRIPTION:\n" +
    (data.description || "No description provided.") + "\n\n" +
    "══════════════════════════════════════\n" +
    "  SUBMISSION META\n" +
    "══════════════════════════════════════\n" +
    "Form Type:   " + (data.form_type   || "Project Inquiry") + "\n" +
    "Source:      " + (data.lead_source || "Website") + "\n" +
    "Submitted:   " + (data.timestamp   || new Date().toISOString()) + "\n\n" +
    "──────────────────────────────────────\n" +
    "Reply directly to this email or contact: " + (data.email || "") + "\n";

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;background:#f9f9f9;padding:24px;border-radius:8px;">' +
    '<div style="background:#1a6fe8;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">' +
    '<h2 style="margin:0;font-size:18px;">🔔 New Project Inquiry</h2>' +
    '<p style="margin:6px 0 0;opacity:.85;font-size:13px;">' + CONFIG.BUSINESS_NAME + ' · ' + new Date().toLocaleString() + '</p>' +
    '</div>' +
    '<div style="background:#fff;padding:24px;border-radius:0 0 8px 8px;border:1px solid #e0e0e0;border-top:none;">' +

    '<h3 style="color:#1a6fe8;font-size:13px;letter-spacing:.08em;text-transform:uppercase;margin:0 0 12px;">Contact Details</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">' +
    tr("Name",    data.full_name) +
    tr("Email",   '<a href="mailto:' + data.email + '">' + data.email + '</a>') +
    tr("Phone",   data.phone) +
    tr("Company", data.company) +
    '</table>' +

    '<h3 style="color:#1a6fe8;font-size:13px;letter-spacing:.08em;text-transform:uppercase;margin:0 0 12px;">Project Details</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">' +
    tr("Service",   data.service) +
    tr("Stage",     data.project_stage) +
    tr("Budget",    data.budget) +
    tr("Timeline",  data.timeline) +
    '</table>' +

    '<h3 style="color:#1a6fe8;font-size:13px;letter-spacing:.08em;text-transform:uppercase;margin:0 0 8px;">Project Description</h3>' +
    '<div style="background:#f4f7ff;border-left:3px solid #1a6fe8;padding:14px 16px;border-radius:4px;font-size:14px;line-height:1.65;color:#333;margin-bottom:20px;">' +
    escHtml(data.description || "No description provided.") +
    '</div>' +

    '<p style="font-size:12px;color:#888;margin:0;">Form Type: ' + (data.form_type || "Project Inquiry") +
    ' &nbsp;|&nbsp; Source: ' + (data.lead_source || "Website") +
    ' &nbsp;|&nbsp; Submitted: ' + (data.timestamp || new Date().toISOString()) + '</p>' +
    '</div></div>';

  MailApp.sendEmail({
    to:       CONFIG.OWNER_EMAIL,
    replyTo:  data.email || CONFIG.OWNER_EMAIL,
    subject:  subject,
    body:     body,
    htmlBody: htmlBody
  });
}


// ── USER CONFIRMATION EMAIL ───────────────────────────────────

function sendUserConfirmationEmail(data) {
  var firstName = (data.full_name || "there").split(" ")[0];

  var subject = "We've received your inquiry — " + CONFIG.BUSINESS_NAME;

  var body =
    "Hi " + firstName + ",\n\n" +
    "Thank you for reaching out to " + CONFIG.BUSINESS_NAME + ".\n\n" +
    "We've received your project inquiry and our engineering team will review it shortly.\n\n" +
    "Here's a summary of what you submitted:\n" +
    "──────────────────────────────────────\n" +
    "Service:  " + (data.service || "—") + "\n" +
    "Stage:    " + (data.project_stage || "—") + "\n" +
    "Budget:   " + (data.budget || "—") + "\n" +
    "Timeline: " + (data.timeline || "—") + "\n" +
    "──────────────────────────────────────\n\n" +
    "What happens next:\n" +
    "1. Our team reviews your submission (usually within a few hours)\n" +
    "2. We respond within 24 hours\n" +
    "3. If your project aligns with our services, we'll schedule a free 30-minute discovery call\n\n" +
    "If you have any urgent questions, feel free to reply to this email.\n\n" +
    "Warm regards,\n" +
    "The " + CONFIG.BUSINESS_NAME + " Team\n" +
    CONFIG.WEBSITE_URL + "\n";

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;background:#f9f9f9;padding:24px;border-radius:8px;">' +
    '<div style="background:#07080d;color:#fff;padding:22px 24px;border-radius:8px 8px 0 0;text-align:center;">' +
    '<h2 style="margin:0;font-size:20px;letter-spacing:-.02em;"><span style="color:#4a93ff;">Collabify</span> <span style="color:#ff6475;">space</span></h2>' +
    '</div>' +
    '<div style="background:#fff;padding:28px 24px;border-radius:0 0 8px 8px;border:1px solid #e0e0e0;border-top:none;">' +
    '<p style="font-size:15px;color:#222;margin:0 0 16px;">Hi <strong>' + escHtml(firstName) + '</strong>,</p>' +
    '<p style="font-size:14px;color:#444;line-height:1.7;margin:0 0 20px;">Thank you for reaching out to <strong>' + CONFIG.BUSINESS_NAME + '</strong>. We\'ve received your project inquiry and our engineering team will review it shortly.</p>' +

    '<div style="background:#f4f7ff;border:1px solid #d0e0ff;border-radius:8px;padding:16px 20px;margin-bottom:22px;">' +
    '<p style="font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#1a6fe8;margin:0 0 10px;">Your Submission Summary</p>' +
    '<table style="width:100%;font-size:13px;color:#444;border-collapse:collapse;">' +
    trUser("Service",  data.service) +
    trUser("Stage",    data.project_stage) +
    trUser("Budget",   data.budget) +
    trUser("Timeline", data.timeline) +
    '</table></div>' +

    '<p style="font-size:13px;font-weight:700;color:#222;margin:0 0 8px;">What happens next:</p>' +
    '<ol style="font-size:13px;color:#555;line-height:1.85;padding-left:20px;margin:0 0 22px;">' +
    '<li>Our team reviews your submission (usually within a few hours)</li>' +
    '<li>We respond within <strong>24 hours</strong></li>' +
    '<li>If your project aligns with our services, we\'ll schedule a free 30-minute discovery call</li>' +
    '</ol>' +

    '<p style="font-size:13px;color:#666;margin:0 0 20px;">If you have any urgent questions, feel free to reply to this email.</p>' +
    '<p style="font-size:14px;color:#222;margin:0;">Warm regards,<br><strong>The ' + CONFIG.BUSINESS_NAME + ' Team</strong></p>' +
    '</div>' +
    '<p style="text-align:center;font-size:11px;color:#aaa;margin-top:16px;">' + CONFIG.WEBSITE_URL + '</p>' +
    '</div>';

  MailApp.sendEmail({
    to:       data.email,
    replyTo:  CONFIG.OWNER_EMAIL,
    subject:  subject,
    body:     body,
    htmlBody: htmlBody
  });
}


// ── UTILITY FUNCTIONS ─────────────────────────────────────────

function tr(label, value) {
  return '<tr>' +
    '<td style="padding:5px 8px 5px 0;color:#888;white-space:nowrap;vertical-align:top;width:90px;">' + label + ':</td>' +
    '<td style="padding:5px 0;color:#222;font-weight:500;">' + (value || '—') + '</td>' +
    '</tr>';
}

function trUser(label, value) {
  return '<tr>' +
    '<td style="padding:3px 8px 3px 0;color:#555;white-space:nowrap;width:80px;">' + label + ':</td>' +
    '<td style="padding:3px 0;color:#222;font-weight:600;">' + escHtml(value || '—') + '</td>' +
    '</tr>';
}

function escHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
