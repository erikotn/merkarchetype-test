/**
 * Merkarchetype Test — Backend
 *
 * Plak deze hele inhoud in Google Apps Script (zie SETUP.md voor stappen).
 * Werkt op een lege Google Sheet — twee tabbladen ('Sessions' en 'Responses')
 * worden automatisch aangemaakt bij de eerste sessie.
 *
 * Admin-toegang (alle sessies overzien):
 *   Stel via Apps Script → Project Settings → Script Properties
 *   een property in met key 'ADMIN_KEY' en een eigen wachtwoord als waarde.
 *   Daarna kun je via de tool met dat wachtwoord alle sessies inzien.
 *
 * API (calls via GET met query params):
 *   ?action=create_session&name=...
 *      → { sessionId, hostToken, name }
 *
 *   ?action=session_info&id=...
 *      → { id, name, status, responseCount }
 *
 *   ?action=submit&id=...&name=...&answers=...
 *      → { ok: true }                                  // mits sessie open is
 *
 *   ?action=results&id=...&host=...
 *      → { sessionName, status, responses: [...] }    // alleen met host_token
 *
 *   ?action=close&id=...&host=...
 *      → { ok: true, status: 'closed' }                // alleen met host_token
 *
 *   ?action=list_sessions&admin=...
 *      → { sessions: [...] }                           // alleen met admin-key
 */

const SESSIONS_TAB = 'Sessions';
const RESPONSES_TAB = 'Responses';

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = (e && e.parameter) ? Object.assign({}, e.parameter) : {};
  if (e && e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      Object.assign(params, body);
    } catch (err) {
      // niet-JSON body → negeren
    }
  }

  let result;
  try {
    switch (params.action) {
      case 'create_session': result = createSession(params); break;
      case 'session_info':   result = sessionInfo(params);   break;
      case 'submit':         result = submitResponse(params); break;
      case 'results':        result = getResults(params);    break;
      case 'close':          result = closeSession(params);  break;
      case 'list_sessions':  result = listSessions(params);  break;
      default:               result = { error: 'Onbekende actie: ' + params.action };
    }
  } catch (err) {
    result = { error: err && err.message ? err.message : String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SESSIONS_TAB) {
      sheet.appendRow(['id', 'name', 'host_token', 'status', 'created_at']);
    } else if (name === RESPONSES_TAB) {
      sheet.appendRow(['session_id', 'participant_name', 'answers', 'submitted_at']);
    }
  }
  return sheet;
}

function generateId() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function generateToken() {
  return Utilities.getUuid().replace(/-/g, '');
}

function createSession(params) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getSheet(SESSIONS_TAB);
    const id = generateId();
    const hostToken = generateToken();
    const name = (params.name || 'Sessie').toString().slice(0, 100);
    sheet.appendRow([id, name, hostToken, 'open', new Date().toISOString()]);
    return { sessionId: id, hostToken: hostToken, name: name };
  } finally {
    lock.releaseLock();
  }
}

function findSession(sessionId) {
  if (!sessionId) return null;
  const sheet = getSheet(SESSIONS_TAB);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId) {
      return {
        row: i + 1,
        id: data[i][0],
        name: data[i][1],
        hostToken: data[i][2],
        status: data[i][3],
        createdAt: data[i][4]
      };
    }
  }
  return null;
}

function countResponses(sessionId) {
  const sheet = getSheet(RESPONSES_TAB);
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId) count++;
  }
  return count;
}

function sessionInfo(params) {
  const session = findSession(params.id);
  if (!session) return { error: 'Sessie niet gevonden' };
  return {
    id: session.id,
    name: session.name,
    status: session.status,
    responseCount: countResponses(params.id)
  };
}

function submitResponse(params) {
  const session = findSession(params.id);
  if (!session) return { error: 'Sessie niet gevonden' };
  if (session.status !== 'open') return { error: 'Sessie is gesloten' };
  if (!params.name || !params.answers) return { error: 'Naam of antwoorden ontbreken' };

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getSheet(RESPONSES_TAB);
    sheet.appendRow([
      params.id,
      params.name.toString().slice(0, 80),
      params.answers.toString().slice(0, 1000),
      new Date().toISOString()
    ]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function getResults(params) {
  const session = findSession(params.id);
  if (!session) return { error: 'Sessie niet gevonden' };
  if (session.hostToken !== params.host) return { error: 'Geen toegang' };

  const sheet = getSheet(RESPONSES_TAB);
  const data = sheet.getDataRange().getValues();
  const responses = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.id) {
      let answers;
      try { answers = JSON.parse(data[i][2]); } catch (e) { answers = {}; }
      responses.push({
        name: data[i][1],
        answers: answers,
        submittedAt: data[i][3]
      });
    }
  }
  return {
    sessionName: session.name,
    status: session.status,
    responses: responses
  };
}

function closeSession(params) {
  const session = findSession(params.id);
  if (!session) return { error: 'Sessie niet gevonden' };
  if (session.hostToken !== params.host) return { error: 'Geen toegang' };

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getSheet(SESSIONS_TAB);
    sheet.getRange(session.row, 4).setValue('closed');
    return { ok: true, status: 'closed' };
  } finally {
    lock.releaseLock();
  }
}

function getAdminKey() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
}

/**
 * Run deze functie eenmalig om je admin-wachtwoord in te stellen.
 *
 * In de Apps Script editor: kies 'eenmaligWachtwoordZetten' uit de
 * dropdown naast het ▶ Run-knopje, klik Run. Er verschijnt een dialoog
 * in je Sheet waarin je het wachtwoord kunt typen. Werkt direct, geen
 * redeploy nodig.
 *
 * Wachtwoord wijzigen? Run 'm gewoon opnieuw — de nieuwe waarde
 * overschrijft de oude. Of wis 'm via Project Settings → Script Properties.
 */
function eenmaligWachtwoordZetten() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Admin-wachtwoord instellen',
    'Type het wachtwoord dat je in de tool wilt gebruiken:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const pwd = response.getResponseText().trim();
  if (!pwd) {
    ui.alert('Geen wachtwoord ingevoerd. Niets opgeslagen.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', pwd);
  ui.alert('Wachtwoord opgeslagen', 'Je kunt nu inloggen via "Beheer" in de tool.', ui.ButtonSet.OK);
}

function listSessions(params) {
  const adminKey = getAdminKey();
  if (!adminKey) return { error: 'Admin-toegang niet geconfigureerd. Stel ADMIN_KEY in via Script Properties.' };
  if (!params.admin || params.admin !== adminKey) return { error: 'Verkeerd wachtwoord' };

  const sheet = getSheet(SESSIONS_TAB);
  const data = sheet.getDataRange().getValues();
  const responsesSheet = getSheet(RESPONSES_TAB);
  const responsesData = responsesSheet.getDataRange().getValues();

  // Tel responses per session in één pass
  const counts = {};
  for (let i = 1; i < responsesData.length; i++) {
    const sid = responsesData[i][0];
    counts[sid] = (counts[sid] || 0) + 1;
  }

  const sessions = [];
  for (let i = 1; i < data.length; i++) {
    sessions.push({
      id: data[i][0],
      name: data[i][1],
      hostToken: data[i][2],
      status: data[i][3],
      createdAt: data[i][4],
      responseCount: counts[data[i][0]] || 0
    });
  }
  // Nieuwste eerst
  sessions.sort((a, b) => {
    const da = new Date(a.createdAt).getTime();
    const db = new Date(b.createdAt).getTime();
    return db - da;
  });
  return { sessions: sessions };
}
