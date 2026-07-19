/**
 * ADEES EXPORTS — FACTORY OS
 * Final Server.gs — Clean, no DASHBOARD dependency
 */

var CONFIG = {
  LIVE_SHEET_ID: '1FLPeuQFPx0nQXRy-16P2-1-e5SjDu7nLE-1ycNZ-IH0',
  DEV_SHEET_ID: '1eHnrG7IWn5PhreW1ywkdhgpzjOzYs6Y53vC4EIxwTvg',
  ENV: 'DEV'
};
var SHEET_ID = CONFIG.ENV === 'DEV' ? CONFIG.DEV_SHEET_ID : CONFIG.LIVE_SHEET_ID;

const ROLES = {
  "ayush@adeesexports.com":   "admin",
  "accounts@adeesexports.com": "accounts",
  "hr@adeesexports.com":      "store",
};

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index').evaluate()
    .setTitle('FactoryOS — Adees Exports' + (CONFIG.ENV === 'DEV' ? ' (DEV)' : ''))
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserInfo() {
  var email = Session.getActiveUser().getEmail();
  var role  = ROLES[email] || 'viewer';
  var first = email.split('@')[0];
  var name  = first.charAt(0).toUpperCase() + first.slice(1);
  return { email:email, role:role, name:name, env:CONFIG.ENV };
}

function safeNum(val) {
  if (val === null || val === undefined) return 0;
  var n = Number(val);
  return isNaN(n) ? 0 : n;
}

function safeStr(val) {
  if (val === null || val === undefined) return '';
  return String(val);
}

function isArtSheet(s) {
  var n = s.getName();
  return n.indexOf('ART-') === 0 && n !== 'ART-TEMPLATE';
}

function getAllData() {
  var result = { ok:true, dash:null, entry:null, wip:null, reqs:null, jobCards:null, user:null };
  try {
    var user = getUserInfo();
    result.user = user;
    Logger.log('[getAllData] user=' + user.email + ' role=' + user.role);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    // dash is only rendered on admin/accounts home. Store home never reads it,
    // so we skip the heavy getDashboardData scan on the floor-staff boot path.
    if (user.role === 'accounts') {
      result.dash = getDashboardData(ss);
      result.entry = getEntryData(null, ss);
      result.reqs = getPendingRequests(ss);
    } else if (user.role === 'admin') {
      result.dash = getDashboardData(ss);
      result.reqs = getPendingRequests(ss);
    } else if (user.role === 'store') {
      result.wip = getWipEntries({}, ss);
      result.jobCards = getJobCards({});
    }
  } catch(e) {
    Logger.log('getAllData error: ' + e.message + ' | stack: ' + e.stack);
    result.error = e.message;
  }
  return result;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
