/**
 * ============================================================================
 * ER. CHETAN KUMAR SARRAF STUDIO — UNIFIED GOOGLE APPS SCRIPT
 * ============================================================================
 * Sheet File: chetan-interior
 * Connected Tabs:
 *   1. "Users"        — User accounts (Google_ID, Name, Email, Photo, First_Login, Last_Login, Login_Count, Role, Password)
 *   2. "Messages"     — Live Google Chat Desk logs + Instant Email Alerts to master.cks009@gmail.com
 *   3. "Reviews"      — Visitor star ratings & client reviews
 *   4. Password Reset — 6-Digit Email Verification Code & Password Reset Facility
 *
 * HOW TO UPDATE IN GOOGLE APPS SCRIPT:
 *   1. Open your "chetan-interior" Google Sheet -> Extensions -> Apps Script.
 *   2. Select all code in Code.gs, delete it, and paste this entire file.
 *   3. Click Save (Ctrl+S).
 *   4. Click Deploy -> Manage deployments -> Edit (pencil icon) ->
 *      Version: "New version" -> Click Deploy. (Your Web App URL stays the same!)
 * ============================================================================
 */

/**
 * ----------------------------------------------------------------------------
 * ONE-CLICK EMAIL AUTHORIZATION TRIGGER (Run this once in Apps Script!)
 * ----------------------------------------------------------------------------
 * 1. At the top toolbar of Apps Script, select "authorizeEmailAndTest" in the function dropdown.
 * 2. Click "Run" (▶).
 * 3. A popup will appear: "Authorization required".
 * 4. Click "Review permissions" -> Select your Gmail (master.cks009@gmail.com).
 * 5. Click "Advanced" -> Click "Go to Untitled project (unsafe)" -> Click "Allow".
 * That's it! Your script will now have permanent permission to send emails and codes!
 */
function authorizeEmailAndTest() {
  MailApp.sendEmail({
    to: 'master.cks009@gmail.com',
    subject: '✓ Er. Chetan Kumar Sarraf Studio — Email Sending Authorized!',
    body: 'Congratulations Er. Chetan! Your Google Apps Script is now permanently authorized to send verification codes, password resets, and chat notifications.'
  });
  Logger.log('Success: Email authorization verified!');
}

var ADMIN_EMAILS = [
  'master.cks009@gmail.com',
  'master.er.cks@outlook.com'
];

var NOTIFY_EMAIL = 'master.cks009@gmail.com';
var EXPECTED_GOOGLE_CLIENT_ID = '619955856662-7d6fpmorphkkq4deri46eq417sh4nrji.apps.googleusercontent.com';

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  } else if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

/**
 * Hash password with SHA-256
 */
function hashPassword(pass) {
  if (!pass) return '';
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pass).trim());
  return Utilities.base64EncodeWebSafe(raw);
}

/**
 * 1. USERS TAB MANAGEMENT (Exact match for user's chetan-interior "Users" sheet)
 * Columns:
 *   A (1): Google_ID
 *   B (2): Name
 *   C (3): Email
 *   D (4): Photo
 *   E (5): First_Login
 *   F (6): Last_Login
 *   G (7): Login_Count
 *   H (8): Role
 *   I (9): Password_Hash
 */
function upsertUserRow(ss, user) {
  var sh = getOrCreateSheet(ss, 'Users', [
    'Google_ID', 'Name', 'Email', 'Photo', 'First_Login', 'Last_Login', 'Login_Count', 'Role', 'Password_Hash'
  ]);
  var nowIso = new Date().toISOString();
  var emailLower = String(user.email || '').trim().toLowerCase();
  var isAdmin = ADMIN_EMAILS.indexOf(emailLower) !== -1;
  var role = isAdmin ? 'ADMIN' : (user.role || 'CLIENT');

  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][2] || '').trim().toLowerCase();
    var rowId = String(data[i][0] || '').trim();
    if ((emailLower && rowEmail === emailLower) || (user.uid && rowId === user.uid)) {
      var count = (parseInt(data[i][6], 10) || 1) + 1; // Col G (7th column)
      var rowNum = i + 1;
      if (user.name) sh.getRange(rowNum, 2).setValue(user.name);
      if (user.picture) sh.getRange(rowNum, 4).setValue(user.picture);
      sh.getRange(rowNum, 6).setValue(nowIso); // Col F: Last_Login
      sh.getRange(rowNum, 7).setValue(count);  // Col G: Login_Count
      sh.getRange(rowNum, 8).setValue(role);   // Col H: Role
      if (user.password) {
        sh.getRange(rowNum, 9).setValue(hashPassword(user.password)); // Col I: Password_Hash
      }
      return {
        uid: rowId || user.uid,
        name: user.name || data[i][1] || 'Client',
        email: emailLower,
        picture: user.picture || data[i][3] || '',
        loginCount: count,
        role: role
      };
    }
  }

  var newUid = user.uid || ('user_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16));
  var passHash = user.password ? hashPassword(user.password) : '';
  sh.appendRow([
    newUid,
    user.name || 'Client',
    emailLower,
    user.picture || '',
    nowIso,
    nowIso,
    1,
    role,
    passHash
  ]);

  return {
    uid: newUid,
    name: user.name || 'Client',
    email: emailLower,
    picture: user.picture || '',
    loginCount: 1,
    role: role
  };
}

/**
 * 2. MESSAGES TAB MANAGEMENT (Exact match for user's chetan-interior "Messages" sheet)
 * Columns:
 *   A: Timestamp
 *   B: Name
 *   C: Email
 *   D: Messages
 *   E: Page
 *   F: AI_Reply
 *   G: Status
 */
function handleChatMessage(ss, payload) {
  var sh = getOrCreateSheet(ss, 'Messages', [
    'Timestamp', 'Name', 'Email', 'Messages', 'Page', 'AI_Reply', 'Status'
  ]);
  var nowIso = new Date().toISOString();
  var name = String(payload.name || 'Studio Visitor').trim();
  var email = String(payload.email || 'visitor@chetankumarsarraf.com.np').trim();
  var msg = String(payload.message || payload.messages || '').trim();
  var page = String(payload.page || 'studio_chat').trim();
  var isSignedIn = payload.is_signed_in === true || payload.is_signed_in === 'true';
  var isOfftopic = payload.is_offtopic === true || payload.is_offtopic === 'true';

  if (!msg) return jsonResponse({ ok: false, error: 'Empty message' });

  var aiReply = isOfftopic ? 'Forwarded to Er. Chetan Kumar Sarraf (Live Desk)' : 'Recorded for Er. Chetan Kumar Sarraf';
  var status = isSignedIn ? (isOfftopic ? 'CLIENT_OFFTOPIC' : 'CLIENT_MESSAGE') : 'VISITOR_MESSAGE';

  // Append directly into the user's 7 columns
  sh.appendRow([
    nowIso,
    name,
    email,
    msg,
    page,
    aiReply,
    status
  ]);

  // Send instant high-priority email notification to master.cks009@gmail.com
  var subjectPrefix = (isSignedIn && isOfftopic) ? '🚨 Direct Client Inquiry: ' : '💬 Studio Chat: ';
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: (email.indexOf('@') !== -1 && email !== 'visitor@chetankumarsarraf.com.np') ? email : NOTIFY_EMAIL,
      subject: subjectPrefix + name + ' (' + email + ')',
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:560px;padding:24px;border:1px solid #e4e4e7;border-radius:10px;background:#ffffff;">' +
        '<div style="font-family:monospace;font-size:11px;color:#d97706;letter-spacing:0.15em;text-transform:uppercase;font-weight:700;">ER. CHETAN KUMAR SARRAF STUDIO DESK</div>' +
        '<h3 style="margin:6px 0 6px;color:#09090b;font-size:18px;">' + (isOfftopic ? 'Direct Client Inquiry from Website' : 'New Studio Chat Message') + '</h3>' +
        '<p style="margin:0 0 16px;color:#71717a;font-size:13px;">' + (isSignedIn ? '✓ Verified Signed-In Client Account' : 'Guest Visitor') + '</p>' +
        '<div style="background:#f4f4f5;padding:16px;border-radius:8px;margin:16px 0;font-size:14px;line-height:1.5;">' +
        '<strong>Client Name:</strong> ' + name + '<br/>' +
        '<strong>Client Email:</strong> ' + email + '<br/>' +
        '<strong>Account Status:</strong> ' + (isSignedIn ? '<span style="color:#16a34a;font-weight:700;">Signed In</span>' : 'Unregistered') + '<br/>' +
        '<strong>Inquiry Type:</strong> ' + (isOfftopic ? '<span style="color:#d97706;font-weight:700;">Custom / Direct Project Inquiry</span>' : 'Standard Inquiry') + '<br/>' +
        '<strong>Time:</strong> ' + nowIso + '<br/><br/>' +
        '<strong>Client Message:</strong><br/>' +
        '<div style="background:#ffffff;padding:14px;border-radius:6px;border:1px solid #e4e4e7;margin-top:6px;color:#18181b;font-size:15px;line-height:1.5;">' +
        msg.replace(/\n/g, '<br/>') +
        '</div>' +
        '</div>' +
        '<div style="margin-top:16px;display:flex;gap:10px;">' +
        '<a href="https://mail.google.com/chat/u/0/#chat/dm/' + encodeURIComponent(email) + '" style="background:#09090b;color:#ffffff;padding:10px 16px;text-decoration:none;font-size:13px;font-weight:700;border-radius:6px;display:inline-block;">Reply on Google Chat</a>' +
        '</div>' +
        '<p style="font-size:12px;color:#a1a1aa;margin-top:16px;">Recorded automatically to your Google Sheet <strong>chetan-interior &rarr; Messages</strong>.</p>' +
        '</div>'
    });
  } catch (err) {}

  return jsonResponse({
    ok: true,
    success: true,
    message: 'Message logged to Messages tab & notification emailed to studio.',
    loggedAt: nowIso
  });
}

/**
 * 3. FORGOT PASSWORD & EMAIL VERIFICATION CODE
 */
function handleForgotPassword(ss, payload) {
  var email = String(payload.email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) {
    return jsonResponse({ ok: false, error: 'Please enter a valid registered email address.' });
  }

  // Accept client-provided code or generate a secure 6-digit verification code
  var code = payload.code ? String(payload.code).trim() : String(Math.floor(100000 + Math.random() * 900000));
  var cache = CacheService.getScriptCache();
  cache.put('cks_reset_' + email, code, 900); // Valid for 15 minutes (900s)

  // Send real email to visitor with verification code
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Er. Chetan Kumar Sarraf Studio — Your 6-Digit Password Reset Code: ' + code,
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:500px;padding:26px;border:1px solid #e4e4e7;border-radius:10px;background:#ffffff;">' +
        '<div style="font-family:monospace;font-size:11px;color:#d97706;letter-spacing:0.15em;text-transform:uppercase;font-weight:700;">ER. CHETAN KUMAR SARRAF STUDIO</div>' +
        '<h2 style="margin:6px 0 10px;color:#09090b;font-size:20px;">Password Reset Verification Code</h2>' +
        '<p style="font-size:14px;color:#52525b;line-height:1.5;">You requested to reset your password for the AI 2D/3D Floor Studio portal. Please enter this 6-digit verification code to confirm your identity:</p>' +
        '<div style="font-size:32px;font-weight:800;letter-spacing:8px;padding:18px 24px;background:#f4f4f5;color:#09090b;text-align:center;margin:20px 0;border-radius:8px;border:1px dashed #d4d4d8;">' +
        code +
        '</div>' +
        '<p style="font-size:12px;color:#71717a;line-height:1.4;">This verification code expires in 15 minutes.<br/>If you did not request a password reset, you can safely ignore this email.</p>' +
        '<hr style="border:none;border-top:1px solid #e4e4e7;margin:20px 0;" />' +
        '<div style="font-size:12px;color:#a1a1aa;text-align:center;">Er. Chetan Kumar Sarraf · Civil Engineer &amp; Spatial Designer · Birgunj, Nepal</div>' +
        '</div>'
    });
  } catch (mailErr) {
    return jsonResponse({ ok: false, error: 'Could not send email: ' + mailErr.toString() });
  }

  return jsonResponse({
    ok: true,
    success: true,
    message: 'Verification code sent to ' + email
  });
}

/**
 * 4. RESET PASSWORD VERIFICATION
 */
function handleResetPassword(ss, payload) {
  var email = String(payload.email || '').trim().toLowerCase();
  var code = String(payload.code || '').trim();
  var newPassword = String(payload.newPassword || payload.password || '').trim();

  if (!email || email.indexOf('@') === -1) {
    return jsonResponse({ ok: false, error: 'Invalid email address.' });
  }
  if (!code || code.length !== 6) {
    return jsonResponse({ ok: false, error: 'Please enter the 6-digit verification code sent to your email.' });
  }
  if (!newPassword || newPassword.length < 6) {
    return jsonResponse({ ok: false, error: 'New password must be at least 6 characters.' });
  }

  var cache = CacheService.getScriptCache();
  var cachedCode = cache.get('cks_reset_' + email);
  if (!cachedCode) {
    return jsonResponse({ ok: false, error: 'Verification code expired or not found. Please request a new code.' });
  }
  if (cachedCode !== code) {
    return jsonResponse({ ok: false, error: 'Incorrect verification code. Please check your email and try again.' });
  }

  // Clear code from cache
  cache.remove('cks_reset_' + email);

  // Update password in Users tab
  var userProfile = upsertUserRow(ss, {
    email: email,
    password: newPassword
  });

  return jsonResponse({
    ok: true,
    success: true,
    message: 'Password reset successful! You can now sign in with your new password.',
    user: userProfile
  });
}

/**
 * 5. GOOGLE SIGN IN VERIFICATION
 */
function handleGoogleLogin(ss, payload) {
  var idToken = payload.idToken || payload.credential || '';
  if (!idToken) {
    // If client already decoded profile or email
    if (payload.email) {
      var profile = upsertUserRow(ss, {
        uid: payload.uid || ('google_' + payload.email.split('@')[0]),
        name: payload.name || payload.email.split('@')[0],
        email: payload.email,
        picture: payload.picture || payload.photo || '',
        role: 'CLIENT'
      });
      return jsonResponse({ ok: true, success: true, user: profile });
    }
    return jsonResponse({ ok: false, error: 'Missing Google credentials.' });
  }

  try {
    var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), {
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) {
      var info = JSON.parse(resp.getContentText());
      if (info && info.email) {
        var profile = upsertUserRow(ss, {
          uid: 'google_' + (info.sub || Utilities.getUuid()),
          name: info.name || info.email.split('@')[0],
          email: info.email,
          picture: info.picture || '',
          role: 'CLIENT'
        });
        return jsonResponse({ ok: true, success: true, user: profile });
      }
    }
  } catch (e) {}

  return jsonResponse({ ok: false, error: 'Could not verify Google ID token.' });
}

/**
 * 6. USER REGISTRATION & LOGIN
 */
function handleUserAuth(ss, payload) {
  var email = String(payload.email || '').trim().toLowerCase();
  var name = String(payload.name || '').trim();
  var pass = String(payload.password || '').trim();
  if (!email || email.indexOf('@') === -1) {
    return jsonResponse({ ok: false, error: 'Valid email required.' });
  }

  var profile = upsertUserRow(ss, {
    name: name,
    email: email,
    password: pass,
    role: 'CLIENT'
  });

  // Send real welcome email to visitor upon registration
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Welcome to Er. Chetan Kumar Sarraf Studio — Account Registered',
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:500px;padding:24px;border:1px solid #e4e4e7;border-radius:10px;background:#ffffff;">' +
        '<div style="font-family:monospace;font-size:11px;color:#d97706;letter-spacing:0.15em;text-transform:uppercase;font-weight:700;">ER. CHETAN KUMAR SARRAF STUDIO</div>' +
        '<h2 style="color:#09090b;margin:8px 0 10px;font-size:20px;">Welcome, ' + (name || 'Client') + '!</h2>' +
        '<p style="font-size:14px;color:#52525b;line-height:1.5;">Your account for Er. Chetan Kumar Sarraf\'s AI 2D/3D Floor Studio has been successfully created.</p>' +
        '<div style="background:#f4f4f5;padding:12px 16px;border-radius:6px;margin:16px 0;font-size:13px;color:#18181b;">' +
        '<strong>Registered Email:</strong> ' + email + '<br/>' +
        '<strong>Account Status:</strong> Active &amp; Cloud-Synced' +
        '</div>' +
        '<p style="font-size:12px;color:#71717a;">You can now log in, design custom floor plans, generate BOQs, and save architectural sessions.</p>' +
        '<hr style="border:none;border-top:1px solid #e4e4e7;margin:18px 0;" />' +
        '<div style="font-size:12px;color:#a1a1aa;text-align:center;">Er. Chetan Kumar Sarraf · Civil Engineer &amp; Spatial Designer · Birgunj, Nepal</div>' +
        '</div>'
    });
  } catch(mailErr) {}

  return jsonResponse({ ok: true, success: true, user: profile });
}

/**
 * 7. CLIENT REVIEW SUBMISSION
 */
function handleReview(ss, data) {
  var sh = getOrCreateSheet(ss, 'Reviews', ['Name', 'Rating', 'City', 'Project Scope', 'Review Comment', 'Date']);
  var name = String(data.name || '').trim();
  var comment = String(data.comment || '').trim();
  if (!name || !comment || /^(\(name\)|name)$/i.test(name)) {
    return jsonResponse({ ok: false, error: 'Empty review data.' });
  }
  sh.appendRow([
    name,
    Number(data.rating) || 5,
    String(data.city || 'Birgunj, Nepal').trim(),
    String(data.scope || 'Residential Interior Design').trim(),
    comment,
    String(data.date || new Date().toISOString().slice(0, 10))
  ]);
  return jsonResponse({ ok: true, success: true, status: 'review_saved' });
}

/**
 * MAIN POST DISPATCHER (Accepts JSON & form data from browser)
 */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = String(data.action || '').toLowerCase();

    if (action === 'chat_message') return handleChatMessage(ss, data);
    if (action === 'forgot_password') return handleForgotPassword(ss, data);
    if (action === 'reset_password') return handleResetPassword(ss, data);
    if (action === 'google_login') return handleGoogleLogin(ss, data);
    if (action === 'user_login' || action === 'register' || action === 'signup') return handleUserAuth(ss, data);
    if (action === 'review') return handleReview(ss, data);

    // If review fields are present by default
    if (data.name && data.comment) {
      return handleReview(ss, data);
    }

    return jsonResponse({ ok: true, status: 'acknowledged', action: action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

/**
 * MAIN GET DISPATCHER (For zero-CORS browser beacons & ping checks)
 */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = String(p.action || 'status').toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'chat_message') return handleChatMessage(ss, p);
  if (action === 'forgot_password') return handleForgotPassword(ss, p);
  if (action === 'reset_password') return handleResetPassword(ss, p);
  if (action === 'user_login' || action === 'register') return handleUserAuth(ss, p);
  if (action === 'review') return handleReview(ss, p);
  if (action === 'ping') {
    return jsonResponse({
      ok: true,
      service: 'Er. Chetan Kumar Sarraf Studio Backend Unified',
      status: 'active',
      clientId: EXPECTED_GOOGLE_CLIENT_ID
    });
  }

  return jsonResponse({ ok: true, status: 'active', client: 'chetan-interior' });
}
