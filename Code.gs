// Google Apps Script — Google Sheet-এ ইউজার ডাটা জমা রাখে (GitHub Pages-এর ব্যাকএন্ড)
var ADMIN_PASSWORD = "QR_Code_Generator@";            // <-- এটি অবশ্যই বদলে নিন
var SITE_NAME = "QR স্টুডিও";                 // ইমেইলে যে নাম দেখাবে
var SHEET_NAME = "Users";
var HEADERS = ["User ID", "Email", "Password", "Status", "Signup Time", "Last Login", "Session Token"];
var EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

function doGet() { return out({ ok: true, msg: "Server চালু আছে" }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents || "{}");
    switch (d.action) {
      case "signup": return out(signup(d));
      case "activate": return out(activate(d));
      case "resend": return out(resend(d));
      case "login": return out(login(d));
      case "verify": return out(verify(d));
      case "logout": return out(logout(d));
      case "forgot": return out(forgot(d));
      case "admin_list": return out(adminList(d));
      case "admin_block": return out(adminBlock(d, true));
      case "admin_unblock": return out(adminBlock(d, false));
      case "admin_set_password": return out(adminSetPassword(d));
      default: return out({ ok: false, error: "অজানা অনুরোধ" });
    }
  } catch (err) {
    return out({ ok: false, error: "সার্ভার সমস্যা: " + err });
  } finally { lock.releaseLock(); }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sh;
}

function now() { return Utilities.formatDate(new Date(), "Asia/Dhaka", "yyyy-MM-dd HH:mm:ss"); }

// সব মান টেক্সট হিসেবে লেখা হয়, যাতে "007123" এর মতো পাসওয়ার্ডের শূন্য হারিয়ে না যায়
function setText(row, col, val) {
  var rg = sheet().getRange(row, col, 1, 1);
  rg.setNumberFormat("@"); rg.setValues([[String(val)]]);
}
function addRow(arr) {
  var sh = sheet(), r = sh.getLastRow() + 1;
  var rg = sh.getRange(r, 1, 1, arr.length);
  rg.setNumberFormat("@"); rg.setValues([arr.map(String)]);
}

function users() {
  var vals = sheet().getDataRange().getValues(), list = [];
  for (var i = 1; i < vals.length; i++) {
    if (!vals[i][1]) continue;
    list.push({ row: i + 1, id: String(vals[i][0]), email: String(vals[i][1]).trim().toLowerCase(),
      pw: String(vals[i][2]), status: String(vals[i][3] || "Active"),
      signup: String(vals[i][4]), last: String(vals[i][5]), token: String(vals[i][6] || "") });
  }
  return list;
}
function blocked(u) { return u.status.trim().toLowerCase() === "blocked"; }
function pending(u) { return u.status.trim().toLowerCase() === "pending"; }
function newToken() { return Utilities.getUuid() + Utilities.getUuid(); }

// অনুমান করা যায় না এমন User ID (যেমন U7K3M9Q) — এটাই ইমেইল যাচাইয়ের কোড
function genId(us) {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  for (var t = 0; t < 50; t++) {
    var id = "U";
    for (var k = 0; k < 6; k++) id += chars.charAt(Math.floor(Math.random() * chars.length));
    var clash = false;
    for (var i = 0; i < us.length; i++) if (us[i].id.toUpperCase() === id) { clash = true; break; }
    if (!clash) return id;
  }
  return "U" + new Date().getTime();
}

function sendId(email, id) {
  MailApp.sendEmail(email, SITE_NAME + " — আপনার User ID",
    "আপনার User ID: " + id + "\n\nএটি কপি করে সাইটের \"অ্যাকাউন্ট চালু করুন\" ঘরে বসান, তাহলে অ্যাকাউন্ট চালু হবে।" +
    "\n\nআপনি সাইন আপ না করে থাকলে এই ইমেইলটি উপেক্ষা করুন।");
}

// সাইন আপ: অ্যাকাউন্ট "Pending" থাকে, User ID শুধু ইমেইলে যায় (স্ক্রিনে দেখানো হয় না)
function signup(d) {
  var email = String(d.email || "").trim().toLowerCase(), pw = String(d.password || "");
  if (!EMAIL_RE.test(email)) return { ok: false, error: "সঠিক ইমেইল দিন।" };
  if (pw.length < 6) return { ok: false, error: "পাসওয়ার্ড কমপক্ষে ৬ অক্ষরের হতে হবে।" };
  var us = users(), ex = null;
  for (var i = 0; i < us.length; i++) if (us[i].email === email) ex = us[i];
  var cache = CacheService.getScriptCache();
  if (ex) {
    if (!pending(ex)) return { ok: false, error: "এই ইমেইল দিয়ে আগেই সাইন আপ করা হয়েছে। লগইন করুন।" };
    if (cache.get("su_" + email)) return { ok: false, error: "একটু আগেই ইমেইল পাঠানো হয়েছে। ১ মিনিট পর আবার চেষ্টা করুন।" };
    try { sendId(email, ex.id); } catch (err) { return { ok: false, error: "ইমেইল পাঠানো যায়নি। ইমেইলটি সঠিক কিনা দেখুন।" }; }
    cache.put("su_" + email, "1", 60);
    setText(ex.row, 3, pw);
    return { ok: true, pending: true, email: email };
  }
  var id = genId(us);
  addRow([id, email, pw, "Pending", now(), "", ""]);
  var rowNo = sheet().getLastRow();
  try { sendId(email, id); }
  catch (err) { sheet().deleteRow(rowNo); return { ok: false, error: "ইমেইল পাঠানো যায়নি। ইমেইলটি সঠিক কিনা দেখুন।" }; }
  cache.put("su_" + email, "1", 60);
  return { ok: true, pending: true, email: email };
}

// ইমেইল থেকে কপি করা User ID বসালে অ্যাকাউন্ট চালু হয় ও লগইন হয়ে যায়
function activate(d) {
  var email = String(d.email || "").trim().toLowerCase(), id = String(d.user_id || "").trim().toUpperCase();
  if (!EMAIL_RE.test(email) || !id) return { ok: false, error: "ইমেইল ও User ID দিন।" };
  var cache = CacheService.getScriptCache(), key = "ak_" + email;
  var tries = parseInt(cache.get(key) || "0", 10);
  if (tries >= 5) return { ok: false, error: "অনেকবার ভুল হয়েছে। ১০ মিনিট পর আবার চেষ্টা করুন।" };
  var us = users();
  for (var i = 0; i < us.length; i++) {
    var u = us[i];
    if (u.email !== email) continue;
    if (blocked(u)) return { ok: false, error: "এই ইমেইলটি ব্লক করা আছে। অ্যাডমিনের সাথে যোগাযোগ করুন।" };
    if (!pending(u)) return { ok: false, error: "অ্যাকাউন্ট আগেই চালু আছে। লগইন করুন।" };
    if (u.id.toUpperCase() !== id) break;
    var token = newToken();
    setText(u.row, 4, "Active"); setText(u.row, 6, now()); setText(u.row, 7, token);
    cache.remove(key);
    return { ok: true, user_id: u.id, email: u.email, token: token };
  }
  cache.put(key, String(tries + 1), 600);
  return { ok: false, error: "User ID মেলেনি। ইমেইলে পাঠানো ID ঠিকভাবে কপি করে বসান।" };
}

// ইমেইল না পেলে আবার পাঠান
function resend(d) {
  var email = String(d.email || "").trim().toLowerCase();
  var generic = { ok: true, msg: "ইমেইল নিবন্ধিত ও চালু হওয়ার অপেক্ষায় থাকলে User ID আবার পাঠানো হয়েছে। Inbox/Spam দেখুন।" };
  if (!EMAIL_RE.test(email)) return { ok: false, error: "সঠিক ইমেইল দিন।" };
  var cache = CacheService.getScriptCache();
  if (cache.get("su_" + email)) return { ok: false, error: "একটু আগেই ইমেইল পাঠানো হয়েছে। ১ মিনিট পর আবার চেষ্টা করুন।" };
  var us = users();
  for (var i = 0; i < us.length; i++) {
    if (us[i].email !== email || !pending(us[i])) continue;
    try { sendId(email, us[i].id); } catch (err) { return { ok: false, error: "ইমেইল পাঠানো যায়নি।" }; }
    cache.put("su_" + email, "1", 60);
    return generic;
  }
  return generic;
}

function login(d) {
  var email = String(d.email || "").trim().toLowerCase(), pw = String(d.password || "");
  var us = users();
  for (var i = 0; i < us.length; i++) {
    var u = us[i];
    if (u.email !== email) continue;
    if (u.pw !== pw) break;
    if (blocked(u)) return { ok: false, error: "এই ইমেইলটি ব্লক করা আছে। অ্যাডমিনের সাথে যোগাযোগ করুন।" };
    if (pending(u)) return { ok: false, pending: true, error: "অ্যাকাউন্ট এখনো চালু হয়নি। ইমেইলে পাঠানো User ID বসিয়ে চালু করুন।" };
    var token = newToken();
    setText(u.row, 6, now()); setText(u.row, 7, token);
    return { ok: true, user_id: u.id, email: u.email, token: token };
  }
  return { ok: false, error: "ইমেইল বা পাসওয়ার্ড ভুল।" };
}

function verify(d) {
  var t = String(d.token || "");
  if (t.length < 20) return { ok: false };
  var us = users();
  for (var i = 0; i < us.length; i++)
    if (us[i].token === t && !blocked(us[i])) return { ok: true, user_id: us[i].id, email: us[i].email };
  return { ok: false };
}

function logout(d) {
  var t = String(d.token || ""), us = users();
  for (var i = 0; i < us.length; i++) if (t && us[i].token === t) setText(us[i].row, 7, "");
  return { ok: true };
}

// ভুলে গেলে: User ID ও পাসওয়ার্ড শুধু নিবন্ধিত ইমেইলের ইনবক্সে পাঠানো হয় (স্ক্রিনে দেখানো হয় না)
function forgot(d) {
  var email = String(d.email || "").trim().toLowerCase();
  var generic = { ok: true, msg: "ইমেইলটি নিবন্ধিত থাকলে User ID ও পাসওয়ার্ড সেই ইমেইলে পাঠানো হয়েছে। Inbox/Spam দেখুন।" };
  if (!EMAIL_RE.test(email)) return { ok: false, error: "সঠিক ইমেইল দিন।" };
  var cache = CacheService.getScriptCache(), key = "fp_" + email;
  if (cache.get(key)) return { ok: false, error: "একটু আগেই অনুরোধ করা হয়েছে। ৫ মিনিট পর আবার চেষ্টা করুন।" };
  var us = users();
  for (var i = 0; i < us.length; i++) {
    if (us[i].email !== email) continue;
    if (blocked(us[i])) return generic;
    cache.put(key, "1", 300);
    MailApp.sendEmail(email, SITE_NAME + " — আপনার লগইন তথ্য",
      "আপনার লগইন তথ্য:\n\nUser ID: " + us[i].id + "\nইমেইল: " + us[i].email +
      "\nপাসওয়ার্ড: " + us[i].pw + "\n\nআপনি অনুরোধ না করলে এই ইমেইলটি উপেক্ষা করুন।");
    return generic;
  }
  return generic;
}

function isAdmin(d) { return String(d.admin_password || "") === ADMIN_PASSWORD; }
var ADMIN_ERR = { ok: false, error: "অ্যাডমিন পাসওয়ার্ড ভুল।" };

function adminList(d) {
  if (!isAdmin(d)) return ADMIN_ERR;
  return { ok: true, users: users().map(function (u) {
    return { id: u.id, email: u.email, password: u.pw, status: u.status, signup: u.signup, last: u.last }; }) };
}

function adminBlock(d, block) {
  if (!isAdmin(d)) return ADMIN_ERR;
  var email = String(d.email || "").trim().toLowerCase(), us = users();
  for (var i = 0; i < us.length; i++) if (us[i].email === email) {
    setText(us[i].row, 4, block ? "Blocked" : "Active");
    if (block) setText(us[i].row, 7, "");
    return { ok: true };
  }
  return { ok: false, error: "ইউজার পাওয়া যায়নি।" };
}

function adminSetPassword(d) {
  if (!isAdmin(d)) return ADMIN_ERR;
  var email = String(d.email || "").trim().toLowerCase(), pw = String(d.new_password || "");
  if (pw.length < 6) return { ok: false, error: "পাসওয়ার্ড কমপক্ষে ৬ অক্ষরের হতে হবে।" };
  var us = users();
  for (var i = 0; i < us.length; i++) if (us[i].email === email) {
    setText(us[i].row, 3, pw); setText(us[i].row, 7, "");   // পুরনো লগইন সেশন বাতিল
    return { ok: true };
  }
  return { ok: false, error: "ইউজার পাওয়া যায়নি।" };
}
