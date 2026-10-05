/*
 * vTiger CRM Auto Update — user database (Google Apps Script)
 *
 * Paste this file into the Apps Script editor of the Google Sheet below
 * (Extensions → Apps Script), run `setup` once, then deploy it as a
 * Web app with "Execute as: Me" and "Who has access: Anyone". Full steps are
 * in README.md → "User database (Google Sheet)".
 *
 * The Sheet itself is never shared or published. The browser only ever talks
 * to this script, which checks passwords here and returns just the signed-in
 * user's own record — never the whole list (except to an Admin, minus hashes).
 *
 * Passwords are stored as salted, peppered, iterated HMAC-SHA256 hashes. The
 * pepper and the session-token secret live in Script Properties (generated on
 * first use), so a copy of the Sheet alone is not enough to check guesses
 * offline. Don't delete those properties: losing PASSWORD_PEPPER invalidates
 * every stored password (everyone would need an Admin reset).
 */

// The Google Sheet that holds the Users and Settings tabs. An id alone grants
// no access — keep the Sheet itself private (not "anyone with the link").
const SPREADSHEET_ID = '1cq0ra2yQLwY3puqF68Sdhb4DZf7EzqnydbXDppwbCOg';
const USERS_SHEET = 'Users';
const SETTINGS_SHEET = 'Settings';
const USER_COLUMNS = ['navneetk', 'Navneet', 'Kumar', 'Admin', 'xh2go2emiDYeNEt4', 'passwordHash', 'passwordSalt', 'passwordHistory', 'updatedAt'];
// Columns added after the first release — created automatically on an existing Users tab.
const AUTO_ADDED_USER_COLUMNS = ['passwordHistory'];
const SETTING_KEYS = ['url', 'module', 'idField', 'timeoutMinutes'];

const HASH_ROUNDS = 1000;
const MIN_PASSWORD_LENGTH = 8;
const PASSWORD_HISTORY_SIZE = 3; // a new password can't match any of the last 3 (current one included)
const TOKEN_TTL_MS = 12 * 60 * 60 * 1000; // a sign-in stays valid for 12 hours
const MAX_FAILED_LOGINS = 5;
const LOCKOUT_SECONDS = 15 * 60;

/* ---------------------------------------------------------------------
 * Entry points
 * ------------------------------------------------------------------- */
function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (_err) {
    return json_({ ok: false, error: 'Invalid request.', code: 'bad_request' });
  }

  const handler = ACTIONS[body.action];
  if (!handler) return json_({ ok: false, error: 'Unknown action.', code: 'bad_request' });

  try {
    return json_(Object.assign({ ok: true }, handler(body)));
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err), code: err.code || 'error' });
  }
}

/* Open the web app URL in a browser to confirm the deployment works and
 * which Sheet it's using (handy if the Users tab "isn't found"). */
function doGet() {
  const ss = spreadsheet_();
  return json_({
    ok: true,
    service: 'vTiger CRM Auto Update user database',
    sheetId: SPREADSHEET_ID,
    tabs: { Users: !!ss.getSheetByName(USERS_SHEET), Settings: !!ss.getSheetByName(SETTINGS_SHEET) },
  });
}

/* Optional: run from the Apps Script editor (select "setup" → Run) to create
 * the tabs up front. The script also creates them on first use. */
function setup() {
  ensureSheets_();
  Logger.log('Setup complete. Add your first Admin row to the "%s" tab (username, role = admin, accessKey).', USERS_SHEET);
}

/* Creates the Users and Settings tabs with their headers if they don't exist yet. */
function ensureSheets_() {
  const ss = spreadsheet_();
  let users = ss.getSheetByName(USERS_SHEET);
  if (!users) users = ss.insertSheet(USERS_SHEET);
  if (users.getLastRow() === 0) {
    users.appendRow(USER_COLUMNS);
    users.setFrozenRows(1);
  }

  let settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) settings = ss.insertSheet(SETTINGS_SHEET);
  if (settings.getLastRow() === 0) {
    settings.appendRow(['key', 'value']);
    SETTING_KEYS.forEach((k) => settings.appendRow([k, '']));
    settings.setFrozenRows(1);
  }

  secret_('PASSWORD_PEPPER');
  secret_('TOKEN_SECRET');
}

/* ---------------------------------------------------------------------
 * Actions
 * ------------------------------------------------------------------- */
const ACTIONS = {
  /* Public: the shared CRM settings (none of these are secret). */
  getSettings() {
    return { settings: readSettings_() };
  },

  /* Public: lets the login screen show "Create Password" for a first-time user. */
  checkUser(body) {
    const user = findUser_(body.username);
    return { exists: !!user, needsPassword: !!user && !user.passwordHash };
  },

  /* Public: verify username + password. A user with no password yet sets it
   * here, but only when the client says it's deliberately creating one
   * (`create: true`, i.e. the Confirm Password field was shown and matched). */
  login(body) {
    const key = normalizeUsername_(body.username);
    const password = String(body.password || '');
    if (!key) throw error_('Please enter your username.', 'bad_request');
    assertNotLocked_(key);

    const user = findUser_(key);
    if (!user) {
      recordFailure_(key);
      throw error_('"' + body.username + '" isn\'t in the registered user list. Ask your admin to add you.', 'not_found');
    }

    let created = false;
    if (!user.passwordHash) {
      if (!body.create) throw error_('Create your password to finish signing in.', 'needs_password');
      withLock_(() => setPassword_(findUser_(key), password));
      created = true;
    } else if (hashPassword_(password, user.passwordSalt) !== user.passwordHash) {
      recordFailure_(key);
      throw error_('Incorrect password.', 'bad_password');
    }

    clearFailures_(key);
    return { token: issueToken_(user.username), user: publicUser_(user, true), created: created };
  },

  /* Signed in: change your own password. Needs the current one; the new one
   * must pass the password policy and not be one of the last 3. */
  changePassword(body) {
    const user = requireUser_(body.token);
    const key = normalizeUsername_(user.username);
    assertNotLocked_(key);
    if (!user.passwordHash || hashPassword_(String(body.currentPassword || ''), user.passwordSalt) !== user.passwordHash) {
      recordFailure_(key);
      throw error_('Your current password is incorrect.', 'bad_password');
    }
    withLock_(() => setPassword_(findUser_(key), String(body.newPassword || '')));
    clearFailures_(key);
    return {};
  },

  /* Public: "Forgot password?" — the user proves who they are with their own
   * vTiger access key (only they and Admins have it) and sets a new password.
   * Wrong keys count toward the same lockout as wrong passwords. */
  resetOwnPassword(body) {
    const key = normalizeUsername_(body.username);
    if (!key) throw error_('Please enter your username.', 'bad_request');
    assertNotLocked_(key);
    const user = findUser_(key);
    const accessKey = String(body.accessKey || '').trim();
    if (!user || !user.accessKey || accessKey !== user.accessKey) {
      recordFailure_(key);
      throw error_('That username and vTiger access key don\'t match.', 'bad_access_key');
    }
    withLock_(() => setPassword_(findUser_(key), String(body.newPassword || '')));
    clearFailures_(key);
    return {};
  },

  /* Signed in: the caller's own fresh record (role/name/access key) — used on
   * page reload and Reconnect so an Admin's changes apply without a new login. */
  me(body) {
    return { user: publicUser_(requireUser_(body.token), true) };
  },

  /* Admin: every user, without password hashes. */
  listUsers(body) {
    requireAdmin_(body.token);
    return { users: readUsers_().map((u) => publicUser_(u, true)) };
  },

  /* Admin: add a user, or update an existing one's name/role/access key.
   * Never touches the password. A blank access key keeps the existing one. */
  saveUser(body) {
    requireAdmin_(body.token);
    const input = body.user || {};
    const username = String(input.username || '').trim();
    if (!username) throw error_('Enter a username.', 'bad_request');

    return withLock_(() => {
      const existing = findUser_(username);
      const accessKey = String(input.accessKey || '').trim() || (existing ? existing.accessKey : '');
      if (!accessKey) throw error_('Enter an access key for a new user.', 'bad_request');

      const fields = {
        username: existing ? existing.username : username,
        firstName: String(input.firstName || '').trim(),
        lastName: String(input.lastName || '').trim(),
        role: normalizeRole_(input.role),
        accessKey: accessKey,
      };
      if (existing) {
        writeUser_(existing.row, fields);
      } else {
        const sheet = usersSheet_();
        sheet.appendRow(USER_COLUMNS.map(() => ''));
        writeUser_(sheet.getLastRow(), fields);
      }
      return { user: publicUser_(findUser_(username), true), created: !existing };
    });
  },

  /* Admin: clear a user's password so they create a new one at next sign-in.
   * The history is kept, so they still can't go back to a recent password. */
  resetPassword(body) {
    requireAdmin_(body.token);
    return withLock_(() => {
      const user = findUser_(body.username);
      if (!user) throw error_('That user no longer exists.', 'not_found');
      writeUser_(user.row, { passwordHash: '', passwordSalt: '', passwordHistory: JSON.stringify(passwordHistory_(user)) });
      clearFailures_(normalizeUsername_(user.username));
      return {};
    });
  },

  /* Admin: delete a user. You can't remove yourself (avoids locking everyone out). */
  removeUser(body) {
    const admin = requireAdmin_(body.token);
    return withLock_(() => {
      const user = findUser_(body.username);
      if (!user) return {};
      if (normalizeUsername_(user.username) === normalizeUsername_(admin.username)) {
        throw error_('You can\'t remove your own account.', 'bad_request');
      }
      usersSheet_().deleteRow(user.row);
      return {};
    });
  },

  /* Admin: save the shared CRM settings for everyone. */
  saveSettings(body) {
    requireAdmin_(body.token);
    const input = body.settings || {};
    return withLock_(() => {
      const sheet = settingsSheet_();
      const values = sheet.getDataRange().getValues();
      SETTING_KEYS.forEach((k) => {
        const value = input[k] === undefined || input[k] === null ? '' : String(input[k]);
        const idx = values.findIndex((r, i) => i > 0 && String(r[0]).trim() === k);
        if (idx > 0) sheet.getRange(idx + 1, 2).setValue(value);
        else sheet.appendRow([k, value]);
      });
      return { settings: readSettings_() };
    });
  },
};

/* ---------------------------------------------------------------------
 * Sheet access
 * ------------------------------------------------------------------- */
function spreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/* Both tabs are created on first use, so a missing tab never blocks sign-in. */
function usersSheet_() {
  return spreadsheet_().getSheetByName(USERS_SHEET) || (ensureSheets_(), spreadsheet_().getSheetByName(USERS_SHEET));
}

function settingsSheet_() {
  return spreadsheet_().getSheetByName(SETTINGS_SHEET) || (ensureSheets_(), spreadsheet_().getSheetByName(SETTINGS_SHEET));
}

/* Column positions come from the header row, so columns can be reordered. */
function userColumnIndex_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map((h) => String(h).trim());
  const index = {};
  USER_COLUMNS.forEach((c) => {
    let i = headers.indexOf(c);
    if (i < 0 && AUTO_ADDED_USER_COLUMNS.indexOf(c) >= 0) {
      i = headers.length;
      sheet.getRange(1, i + 1).setValue(c);
      headers.push(c);
    }
    if (i < 0) throw error_('The "' + USERS_SHEET + '" tab is missing the "' + c + '" column.', 'misconfigured');
    index[c] = i;
  });
  return index;
}

function readUsers_() {
  const sheet = usersSheet_();
  const index = userColumnIndex_(sheet);
  const values = sheet.getDataRange().getValues();
  const users = [];
  for (let r = 1; r < values.length; r += 1) {
    const username = String(values[r][index.username] || '').trim();
    if (!username) continue;
    const user = { row: r + 1 };
    USER_COLUMNS.forEach((c) => { user[c] = String(values[r][index[c]] === undefined ? '' : values[r][index[c]]).trim(); });
    user.username = username;
    user.role = normalizeRole_(user.role);
    users.push(user);
  }
  return users;
}

function findUser_(username) {
  const key = normalizeUsername_(username);
  if (!key) return null;
  return readUsers_().find((u) => normalizeUsername_(u.username) === key) || null;
}

function writeUser_(row, fields) {
  const sheet = usersSheet_();
  const index = userColumnIndex_(sheet);
  fields.updatedAt = new Date().toISOString();
  Object.keys(fields).forEach((c) => {
    if (index[c] === undefined) return;
    // Plain text format stops Sheets from reinterpreting keys/hashes as numbers or dates.
    sheet.getRange(row, index[c] + 1).setNumberFormat('@').setValue(fields[c]);
  });
}

function readSettings_() {
  const settings = {};
  const sheet = spreadsheet_().getSheetByName(SETTINGS_SHEET);
  if (!sheet) return settings;
  sheet.getDataRange().getValues().slice(1).forEach((r) => {
    const k = String(r[0]).trim();
    const v = String(r[1] === undefined ? '' : r[1]).trim();
    if (SETTING_KEYS.indexOf(k) >= 0 && v !== '') settings[k] = k === 'timeoutMinutes' ? Number(v) || 0 : v;
  });
  return settings;
}

/* ---------------------------------------------------------------------
 * Auth helpers
 * ------------------------------------------------------------------- */
function publicUser_(user, includeAccessKey) {
  const out = {
    username: user.username,
    firstName: user.firstName,
    lastName: user.lastName,
    role: user.role,
    hasPassword: !!user.passwordHash,
  };
  if (includeAccessKey) out.accessKey = user.accessKey;
  return out;
}

function normalizeUsername_(u) {
  return String(u || '').trim().toLowerCase();
}

function normalizeRole_(role) {
  return /admin/i.test(String(role || '')) ? 'admin' : 'member';
}

function hashPassword_(password, salt) {
  const pepper = Utilities.newBlob(secret_('PASSWORD_PEPPER')).getBytes();
  let digest = Utilities.computeHmacSha256Signature(Utilities.newBlob(salt + ':' + password).getBytes(), pepper);
  for (let i = 1; i < HASH_ROUNDS; i += 1) {
    digest = Utilities.computeHmacSha256Signature(digest, pepper);
  }
  return Utilities.base64Encode(digest);
}

/* ---------------------------------------------------------------------
 * Password policy — mirrored in js/app.js (PASSWORD_RULES) for the live
 * checklist, but this is the copy that's actually enforced.
 * ------------------------------------------------------------------- */
function passwordProblems_(password) {
  const problems = [];
  if (password.length < MIN_PASSWORD_LENGTH) problems.push('at least ' + MIN_PASSWORD_LENGTH + ' characters');
  if (!/[A-Z]/.test(password)) problems.push('an uppercase letter');
  if (!/[a-z]/.test(password)) problems.push('a lowercase letter');
  if (!/[0-9]/.test(password)) problems.push('a number');
  if (!/[^A-Za-z0-9]/.test(password)) problems.push('a special character');
  if (/(.)\1\1/.test(password)) problems.push('no character repeated 3 or more times in a row (like "111" or "aaa")');
  if (hasSequentialRun_(password)) problems.push('no 3 or more sequential characters (like "123", "abc" or "321")');
  return problems;
}

/* True if 3+ consecutive digits or letters run up or down by one: 123, 987, abc, CBA. */
function hasSequentialRun_(password) {
  const s = password.toLowerCase();
  for (let i = 0; i + 2 < s.length; i += 1) {
    if (!/^(?:[0-9]{3}|[a-z]{3})$/.test(s.substr(i, 3))) continue;
    const step = s.charCodeAt(i + 1) - s.charCodeAt(i);
    if (Math.abs(step) === 1 && s.charCodeAt(i + 2) - s.charCodeAt(i + 1) === step) return true;
  }
  return false;
}

/* The user's last PASSWORD_HISTORY_SIZE passwords as [{ s: salt, h: hash }],
 * newest first, including the current one. */
function passwordHistory_(user) {
  let list = [];
  try { list = JSON.parse(user.passwordHistory || '[]'); } catch (_err) { /* treat as empty */ }
  if (!Array.isArray(list)) list = [];
  // Rows created before history existed only have the current hash.
  if (user.passwordHash && !list.some((e) => e && e.h === user.passwordHash)) {
    list.unshift({ s: user.passwordSalt, h: user.passwordHash });
  }
  return list.filter((e) => e && e.s && e.h).slice(0, PASSWORD_HISTORY_SIZE);
}

/* Validates and stores a new password for `user`. Call inside withLock_. */
function setPassword_(user, password) {
  if (!user) throw error_('That user no longer exists.', 'not_found');
  const problems = passwordProblems_(password);
  if (problems.length) throw error_('Your password needs ' + problems.join(', ') + '.', 'weak_password');

  const history = passwordHistory_(user);
  if (history.some((e) => hashPassword_(password, e.s) === e.h)) {
    throw error_('You can\'t reuse any of your last ' + PASSWORD_HISTORY_SIZE + ' passwords. Choose a new one.', 'password_reused');
  }

  const salt = Utilities.getUuid();
  const hash = hashPassword_(password, salt);
  writeUser_(user.row, {
    passwordHash: hash,
    passwordSalt: salt,
    passwordHistory: JSON.stringify([{ s: salt, h: hash }].concat(history).slice(0, PASSWORD_HISTORY_SIZE)),
  });
}

/* Token = base64(payload) + "." + HMAC signature. It only carries the username
 * and expiry; the role is always re-read from the Sheet on every request. */
function issueToken_(username) {
  const payload = Utilities.base64EncodeWebSafe(JSON.stringify({ u: username, exp: Date.now() + TOKEN_TTL_MS }));
  return payload + '.' + sign_(payload);
}

function requireUser_(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || sign_(parts[0]) !== parts[1]) {
    throw error_('Your session is invalid — please sign in again.', 'unauthorized');
  }
  const data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  if (!data.exp || data.exp < Date.now()) throw error_('Your session expired — please sign in again.', 'unauthorized');
  const user = findUser_(data.u);
  if (!user) throw error_('Your user entry was removed. Contact your admin.', 'unauthorized');
  return user;
}

function requireAdmin_(token) {
  const user = requireUser_(token);
  if (user.role !== 'admin') throw error_('Only Admins can do that.', 'forbidden');
  return user;
}

function sign_(payload) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_('TOKEN_SECRET')));
}

/* Brute-force protection: too many wrong passwords locks that username for a while. */
function assertNotLocked_(key) {
  const failures = Number(CacheService.getScriptCache().get('fail:' + key) || 0);
  if (failures >= MAX_FAILED_LOGINS) {
    throw error_('Too many failed sign-in attempts. Try again in ' + Math.round(LOCKOUT_SECONDS / 60) + ' minutes.', 'locked');
  }
}

function recordFailure_(key) {
  const cache = CacheService.getScriptCache();
  const failures = Number(cache.get('fail:' + key) || 0) + 1;
  cache.put('fail:' + key, String(failures), LOCKOUT_SECONDS);
}

function clearFailures_(key) {
  CacheService.getScriptCache().remove('fail:' + key);
}

/* ---------------------------------------------------------------------
 * Misc
 * ------------------------------------------------------------------- */
function secret_(name) {
  const props = PropertiesService.getScriptProperties();
  let value = props.getProperty(name);
  if (!value) {
    value = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(name, value);
  }
  return value;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function error_(message, code) {
  const err = new Error(message);
  err.code = code;
  return err;
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
