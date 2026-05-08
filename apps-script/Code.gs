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
const ARCHETYPES_TAB = 'Archetypes';
const QUESTIONS_TAB = 'Questions';
const MAPPINGS_TAB = 'Mappings';
const COMBINATIONS_TAB = 'Combinations';

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
      case 'get_config':     result = getConfig();           break;
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

// ─── CONTENT (Archetypes / Questions / Mappings) ─────────────────

function readSheetAsObjects(tabName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    let hasData = false;
    for (let j = 0; j < headers.length; j++) {
      const v = data[i][j];
      row[headers[j]] = v;
      if (v !== '' && v !== null && v !== undefined) hasData = true;
    }
    if (hasData) rows.push(row);
  }
  return rows;
}

function readArchetypes() {
  return readSheetAsObjects(ARCHETYPES_TAB);
}

function readQuestions() {
  const rows = readSheetAsObjects(QUESTIONS_TAB);
  if (!rows) return null;
  return rows.map(r => ({
    id: parseInt(r.id, 10),
    left: r.left,
    right: r.right
  })).filter(q => !isNaN(q.id) && q.id > 0);
}

function readMappings() {
  const rows = readSheetAsObjects(MAPPINGS_TAB);
  if (!rows) return null;
  // Naar nested object: { qid: { answer: { archetype: points } } }
  const out = {};
  rows.forEach(r => {
    const qid = parseInt(r.question_id, 10);
    const ans = parseInt(r.answer, 10);
    const arch = String(r.archetype || '').trim();
    const pts = parseInt(r.points, 10);
    if (isNaN(qid) || isNaN(ans) || !arch || isNaN(pts)) return;
    if (!out[qid]) out[qid] = {};
    if (!out[qid][ans]) out[qid][ans] = {};
    out[qid][ans][arch] = pts;
  });
  return out;
}

function readCombinations() {
  const rows = readSheetAsObjects(COMBINATIONS_TAB);
  if (!rows) return null;
  // Naar nested object: { primary: { secondary: { description, type } } }
  const out = {};
  rows.forEach(r => {
    const p = String(r.primary || '').trim();
    const s = String(r.secondary || '').trim();
    const d = String(r.description || '').trim();
    const t = String(r.type || '').trim();
    if (!p || !s || !d) return;
    if (!out[p]) out[p] = {};
    out[p][s] = { description: d, type: t };
  });
  return out;
}

function getConfig() {
  return {
    archetypes: readArchetypes(),
    questions: readQuestions(),
    mappings: readMappings(),
    combinations: readCombinations()
  };
}

/**
 * Run deze functie eenmalig om de Archetypes, Questions en Mappings
 * tabs aan te maken en te vullen met de huidige content. Daarna kun je
 * in die tabs alles vrij bewerken — de tool leest live wat je daar zet.
 *
 * Bestaande tabs worden NIET overschreven (veiligheid). Wil je opnieuw
 * vullen? Hernoem of verwijder de tabs eerst.
 */
function eenmaligContentTabsAanmaken() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const existing = [ARCHETYPES_TAB, QUESTIONS_TAB, MAPPINGS_TAB, COMBINATIONS_TAB]
    .filter(t => ss.getSheetByName(t));
  if (existing.length > 0) {
    ui.alert(
      'Tabs bestaan al',
      'Deze tabs bestaan al: ' + existing.join(', ') +
      '\n\nVerwijder of hernoem ze eerst als je opnieuw wilt initialiseren.',
      ui.ButtonSet.OK
    );
    return;
  }

  const seed = SEED_DATA();
  populateSheet(ss, ARCHETYPES_TAB, seed.archetypes);
  populateSheet(ss, QUESTIONS_TAB, seed.questions);
  populateSheet(ss, MAPPINGS_TAB, seed.mappings);
  populateSheet(ss, COMBINATIONS_TAB, seed.combinations);

  ui.alert(
    'Klaar',
    'Vier tabs zijn aangemaakt en gevuld:\n\n' +
    '• Archetypes — namen, kleuren en beschrijvingen\n' +
    '• Questions — de 30 stellingen\n' +
    '• Mappings — punten per (vraag, antwoord, archetype)\n' +
    '• Combinations — eigen merkpersoonlijkheid-tekst per top-2-combinatie (begint met 1 voorbeeld)\n\n' +
    'Bewerk gerust — de tool leest deze tabs live.',
    ui.ButtonSet.OK
  );
}

function populateSheet(ss, tabName, rows) {
  const sheet = ss.insertSheet(tabName);
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight('bold')
    .setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
  // Auto-resize kolommen
  for (let c = 1; c <= rows[0].length; c++) {
    sheet.autoResizeColumn(c);
  }
}

/**
 * Voegt 30 voorgeschreven combinatie-teksten toe aan de Combinations-tab.
 *
 * Werkt in beide situaties:
 * - Combinations-tab bestaat al → rijen worden onder de bestaande gezet.
 *   Bestaande paren blijven ongemoeid; alleen paren die er nog niet staan
 *   (in beide richtingen) worden toegevoegd.
 * - Combinations-tab bestaat nog niet → wordt aangemaakt en gevuld.
 *
 * Run één keer. Erna kun je in de Sheet zelf bewerken/uitbreiden.
 */
/**
 * Vervangt de hele Combinations-tab met de 132 directionele combinaties.
 *
 * - Combinations-tab bestaat nog niet → wordt aangemaakt en gevuld.
 * - Bestaat wel → wordt geleegd en opnieuw gevuld (ná bevestiging).
 *
 * Belangrijk: alle handmatige aanpassingen aan de Combinations-tab gaan
 * verloren. Run alleen als je dat OK vindt.
 *
 * Combinaties zijn richtingsspecifiek: 'innocent + sage' (Onschuldige
 * primair) is een andere tekst dan 'sage + innocent' (Wijze primair).
 */
function vervangAlleCombinaties() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newRows = COMBINATIONS_DATA;
  const dataRows = newRows.slice(1);

  let sheet = ss.getSheetByName(COMBINATIONS_TAB);
  if (sheet) {
    const existingData = sheet.getDataRange().getValues();
    const hasContent = existingData.length > 1;
    if (hasContent) {
      const resp = ui.alert(
        'Combinations-tab vervangen?',
        'De huidige Combinations-tab bevat ' + (existingData.length - 1) + ' rijen.\n\n' +
        'Deze worden allemaal vervangen door de 132 voorgeschreven directionele teksten. ' +
        'Eigen aanpassingen gaan verloren.\n\nDoorgaan?',
        ui.ButtonSet.YES_NO
      );
      if (resp !== ui.Button.YES) return;
    }
    // Wis bestaand
    sheet.clear();
  } else {
    sheet = ss.insertSheet(COMBINATIONS_TAB);
  }

  // Vul met de 132 rijen + header
  sheet.getRange(1, 1, newRows.length, newRows[0].length).setValues(newRows);
  sheet.getRange(1, 1, 1, newRows[0].length)
    .setFontWeight('bold')
    .setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.autoResizeColumn(3); // type
  sheet.setColumnWidth(4, 800); // description (lang)

  ui.alert(
    'Klaar',
    dataRows.length + ' directionele combinatie-teksten staan nu in de Combinations-tab.\n\n' +
    'Bewerk gerust in de Sheet — de tool leest live (cache 1 uur).',
    ui.ButtonSet.OK
  );
}

function SEED_DATA() {
  return {
    archetypes: JSON.parse(ARCHETYPES_SEED),
    questions: JSON.parse(QUESTIONS_SEED),
    mappings: JSON.parse(MAPPINGS_SEED),
    combinations: COMBINATIONS_DATA
  };
}

const ARCHETYPES_SEED = '[["key", "name", "emoji", "subtitle", "color", "promise", "description", "communication", "traits", "brands"], ["innocent", "De Onschuldige", "🌼", "puur, optimistisch, eenvoudig", "#7AC74F", "Het komt allemaal goed.", "De Onschuldige staat voor puurheid, optimisme en eenvoud. Dit archetype gelooft in het goede en wil zowel zichzelf als anderen gelukkig maken. Het is een betrouwbaar, eerlijk en nostalgisch merk dat veiligheid en positiviteit uitstraalt.", "De Onschuldige communiceert op een eenvoudige, heldere en positieve manier. De toon is optimistisch en geruststellend, zonder cynisme of complexiteit. Beeldgebruik is clean, licht en vaak nostalgisch. Boodschappen zijn rechtdoorzee en bieden troost en zekerheid. Humor is vriendelijk en onschuldig. Dit merk spreekt over waarden als eerlijkheid, betrouwbaarheid en traditionele kwaliteit.", "Positiviteit, vertrouwen, eenvoud, nostalgie", "Campina, Friesche Vlag, Verkade, Dove, Coca-Cola"], ["explorer", "De Ontdekker", "🧭", "vrijheid, avontuur, ontdekking", "#9CCC65", "Ontdek jezelf en de wereld.", "De Ontdekker staat voor vrijheid, avontuur en ontdekking. Dit archetype wil grenzen verleggen, nieuwe ervaringen opdoen en authentiek zijn. Het is onafhankelijk, moedig en pionierend, altijd op zoek naar wat er verder nog te ontdekken valt.", "De Ontdekker communiceert inspirerend en bevrijdend. De toon is avontuurlijk en nodigt uit tot actie. Beelden tonen wijde landschappen, reizen en grenzeloosheid. De boodschap gaat over zelfontdekking, individualiteit en het verleggen van grenzen. Dit merk spreekt mensen aan die zich niet willen laten beperken en hun eigen pad willen bewandelen. Verhalen gaan over expedities, ontdekkingen en persoonlijke groei.", "Avontuur, autonomie, authenticiteit, ontdekking", "KLM, ANWB, Staatsloterij, The North Face, Jeep"], ["sage", "De Wijze", "🦉", "kennis, expertise, betrouwbaarheid", "#388E3C", "De waarheid zal je vrijmaken.", "De Wijze staat voor kennis, waarheid en intelligentie. Dit archetype streeft ernaar de wereld te begrijpen en die kennis te delen met anderen. Het is een expert, een betrouwbare bron van informatie die analytisch en doordacht te werk gaat.", "De Wijze communiceert helder, informatief en onderbouwd. De toon is educatief en gezaghebbend, maar niet betuttelend. Feiten, onderzoek en expertise staan centraal. De communicatie daagt uit om na te denken en nodigt uit tot leren. Beeldmateriaal is vaak strak en professioneel. Jargon is toegestaan als het relevant is. Dit merk positioneert zich als thought leader en betrouwbare kennisbron.", "Wijsheid, expertise, analyse, waarheid", "NOS, NU.nl, Douwe Egberts, BBC, Google"], ["hero", "De Held", "🦸", "moed, prestatie, overwinning", "#E53935", "Waar een wil is, is een weg.", "De Held staat voor moed, prestatie en overwinning. Dit archetype wil impact maken en uitdagingen overwinnen. Het is dapper, sterk en gedreven naar uitmuntendheid.", "De Held communiceert motiverend, krachtig en doelgericht. De toon is inspirerend en roept op tot actie en prestatie. Beelden tonen overwinning, doorzettingsvermogen en succes. De boodschap is dat je alles kunt bereiken als je maar hard genoeg werkt. Dit merk daagt mensen uit het beste uit zichzelf te halen. Verhalen gaan over overwinnen van obstakels, competitie en triomf. De communicatie is energiek en vol zelfvertrouwen.", "Moed, doorzettingsvermogen, prestatie, kracht", "Jumbo, Rabobank, Heineken, Nike, Adidas, FedEx"], ["outlaw", "De Rebel", "🔥", "regelbreker, anders, disruptief", "#F57C00", "De regels zijn er om gebroken te worden.", "De Rebel staat voor rebellie, bevrijding en revolutie. Dit archetype daagt de status quo uit en durft te provoceren. Het is radicaal, disruptief en authentiek rebels.", "De Rebel communiceert provocerend, direct en zonder compromissen. De toon is rebels en uitdagend, soms zelfs confronterend. Beeldgebruik is ruig, edgy en onconventioneel. Dit merk spreekt taboes aan en kiest bewust positie. De communicatie roept op tot actie en verandering. Humor kan cynisch of sarcastisch zijn. Dit merk trekt mensen aan die zich niet conformeren en anders durven te zijn.", "Rebellie, bevrijding, revolutie, provocatie", "G-Star RAW, Tony\'s Chocolonely, Transavia, Harley-Davidson, Diesel, Virgin"], ["magician", "De Magiër", "✨", "transformatie, visie, wonder", "#FB8C00", "Alles is mogelijk.", "De Magiër staat voor transformatie, visie en het waar maken van dromen. Dit archetype inspireert door ogenschijnlijk onmogelijke dingen mogelijk te maken. Het is charismatisch, visionair en transformerend.", "De Magiër communiceert mysterieus, inspirerend en visionair. De toon is magisch en belooft transformatie. Beelden zijn vaak sprookjesachtig, verbeeldingsvol of toekomstgericht. De communicatie gaat over dromen die werkelijkheid worden, over magie en wonder. Dit merk toont hoe het het leven van mensen kan transformeren. Verhalen zijn meeslepend en emotioneel. De boodschap is dat met dit merk alles binnen bereik ligt.", "Transformatie, visie, inspiratie, innovatie", "Efteling, Philips, Disney, Tesla"], ["regular", "De Bondgenoot", "🤝", "erbij horen, toegankelijk, realisme", "#7E57C2", "Wij begrijpen je.", "De Bondgenoot staat voor verbinding, eerlijkheid en toegankelijkheid. Dit archetype wil erbij horen en echte connectie maken met anderen. Het is betrouwbaar, down-to-earth en herkenbaar.", "De Bondgenoot communiceert toegankelijk, herkenbaar en zonder poespas. De toon is vriendelijk, eerlijk en praktisch. Beelden tonen gewone mensen in alledaagse situaties. De communicatie is realistisch en zonder pretentie. Dit merk spreekt de taal van de doelgroep en toont begrip voor hun uitdagingen. Verhalen gaan over herkenbare situaties en praktische oplossingen. De boodschap is: wij zijn net als jij en staan aan jouw kant.", "Toegankelijkheid, eerlijkheid, realisme, verbinding", "Albert Heijn, HEMA, Gamma, Kruidvat, IKEA, Levi\'s, eBay"], ["lover", "De Verleider", "🌹", "intimiteit, plezier, sensualiteit", "#EC407A", "Geniet van het moment.", "De Verleider staat voor passie, intimiteit en schoonheid. Dit archetype streeft naar nabijheid, genot en betekenisvolle ervaringen. Het is sensueel, warm en gepassioneerd.", "De Verleider communiceert sensueel, emotioneel en intiem. De toon is warm, verleidelijk en persoonlijk. Beelden zijn esthetisch, romantisch en appelleren aan de zintuigen. De communicatie gaat over genieten, verwennen en het ervaren van schoonheid. Dit merk spreekt over kwaliteitstijd, zintuiglijke ervaringen en het belang van intimiteit en verbinding. Verhalen zijn emotioneel geladen en persoonlijk. De boodschap is dat je het verdient om jezelf te verwennen.", "Passie, intimiteit, schoonheid, sensualiteit", "Rituals, De Bijenkorf, Spa, Godiva, Chanel, Alfa Romeo"], ["jester", "De Joker", "🎭", "plezier, humor, luchtigheid", "#AB47BC", "Het leven is een feest.", "De Joker staat voor plezier, humor en luchtigheid. Dit archetype wil het leven vieren en anderen aan het lachen maken. Het is speels, optimistisch en spontaan.", "De Joker communiceert humoristisch, energiek en ongedwongen. De toon is lichtvoetig, grappig en entertainend. Beelden zijn kleurrijk, speels en onverwacht. De communicatie durft gek te doen en neemt zichzelf niet te serieus. Dit merk gebruikt humor, woordgrappen en onverwachte wendingen. Verhalen zijn vermakelijk en maken mensen aan het lachen. De boodschap is dat het leven te kort is om je zorgen te maken en dat je vooral moet genieten.", "Humor, plezier, spontaniteit, luchtigheid", "Coolblue, Ben, bol.com, Old Spice, M&M\'s, Skittles"], ["caregiver", "De Beschermer", "🛡️", "zorg, bescherming, medeleven", "#5C6BC0", "Ik zorg voor je.", "De Beschermer staat voor zorg, bescherming en compassie. Dit archetype wil anderen helpen en beschermen. Het is warm, ondersteunend en altruïstisch.", "De Beschermer communiceert empathisch, geruststellend en ondersteunend. De toon is zorgzaam, warm en betrokken. Beelden tonen zorg, bescherming en menselijke verbinding. De communicatie draait om het welzijn van de ander en toont begrip voor zorgen en behoeften. Dit merk stelt zich dienstbaar op en spreekt over verantwoordelijkheid en toewijding. Verhalen gaan over het verschil maken in iemands leven en er zijn wanneer het nodig is. De boodschap is dat je erop kunt vertrouwen dat dit merk voor je zorgt.", "Zorg, bescherming, generositeit, empathie", "Zilveren Kruis, CZ, Rode Kruis, VGZ, Unicef, Volvo, Johnson & Johnson"], ["ruler", "De Leider", "👑", "controle, leiderschap, status", "#1E88E5", "Macht creëert mogelijkheden.", "De Leider staat voor controle, leiderschap en succes. Dit archetype streeft naar het creëren van welvaart en stabiliteit. Het is dominant, georganiseerd en exclusief.", "De Leider communiceert zelfverzekerd, autoritair en prestigieus. De toon is formeel, krachtig en gebiedend. Beelden tonen luxe, succes en exclusiviteit. De communicatie gaat over leiderschap, controle en het bereiken van de top. Dit merk positioneert zich als marktleider en statusmerk. Verhalen gaan over excellentie, erfgoed en het bij de elite horen. De boodschap is dat dit merk voor winners is, voor mensen die het beste verdienen en het beste willen.", "Leiderschap, controle, status, stabiliteit", "ABN AMRO, ING, Nationale Nederlanden, Mercedes-Benz, Rolex, Microsoft"], ["creator", "De Creator", "🎨", "innovatie, creativiteit, verbeelding", "#26C6DA", "Als je het kunt bedenken, kun je het maken.", "De Creator staat voor creativiteit, innovatie en verbeelding. Dit archetype wil iets waardevols en blijvends creëren. Het is creatief, innovatief en authentiek non-conform.", "De Creator communiceert inspirerend, origineel en kunstzinnig. De toon is creatief, innovatief en soms onconventioneel. Beelden zijn vaak artistiek, vernieuwend of design-gedreven. De communicatie gaat over zelfexpressie, creativiteit en het benutten van je verbeeldingskracht. Dit merk moedigt aan om dingen te maken, te ontwerpen en jezelf uit te drukken. Verhalen gaan over het creatieve proces, innovatie en het tot leven brengen van ideeën. De boodschap is dat iedereen creatief kan zijn en dat dit merk de tools biedt om je visie te realiseren.", "Creativiteit, innovatie, verbeelding, expressie", "ASML, Coolblue, Apple, Lego, Adobe"]]';

const QUESTIONS_SEED = '[["id","left","right"],[1,"We zijn speels en spontaan","We zijn serieus en verantwoordelijk"],[2,"We streven naar vrijheid en avontuur","We streven naar veiligheid en zekerheid"],[3,"Onze toon is vriendelijk en zorgzaam","Onze toon is sterk en beschermend"],[4,"We willen de wereld verbeteren","We willen het leven gemakkelijker maken"],[5,"We staan voor authenticiteit en echtheid","We staan voor succes en prestige"],[6,"We inspireren door verbeeldingskracht","We inspireren door logica en efficiëntie"],[7,"We zoeken verandering en revolutie","We zoeken stabiliteit en continuïteit"],[8,"Onze communicatie is warm en menselijk","Onze communicatie is cool en professioneel"],[9,"We richten ons op gemeenschap en verbinding","We richten ons op prestatie en status"],[10,"We helpen mensen hun potentieel te zien","We helpen mensen zich veilig te voelen"],[11,"We geloven in dromen en magie","We geloven in feiten en realiteit"],[12,"We willen dat klanten zich vrij voelen","We willen dat klanten zich beschermd weten"],[13,"We spreken als een vriend","We spreken als een expert"],[14,"We moedigen avontuur en ontdekking aan","We moedigen orde en structuur aan"],[15,"We brengen originaliteit","We brengen betrouwbaarheid"],[16,"We streven naar perfectie","We streven naar authenticiteit"],[17,"We staan voor discipline en beheersing","We staan voor passie en emotie"],[18,"We willen anderen inspireren","We willen anderen dienen"],[19,"We waarderen speelsheid","We waarderen wijsheid"],[20,"We geloven in rebellie","We geloven in traditie"],[21,"Onze klanten zien ons als een gids","Onze klanten zien ons als een bondgenoot"],[22,"We zijn visionair en idealistisch","We zijn praktisch en realistisch"],[23,"We streven naar groei en zelfontplooiing","We streven naar comfort en voorspelbaarheid"],[24,"We spreken met humor","We spreken met ernst"],[25,"We positioneren ons als uitdager","We positioneren ons als leider"],[26,"We draaien om emotie","We draaien om ratio"],[27,"We zijn verleidelijk en aantrekkelijk","We zijn nuchter en functioneel"],[28,"We willen anderen verrassen","We willen anderen geruststellen"],[29,"We bouwen ons merk op intuïtie","We bouwen ons merk op analyse"],[30,"We streven naar vrijheid van expressie","We streven naar controle en orde"]]';

const MAPPINGS_SEED = '[["question_id","answer","archetype","points"],[1,1,"jester",2],[1,1,"explorer",1],[1,2,"jester",1],[1,2,"explorer",1],[1,4,"ruler",1],[1,4,"caregiver",1],[1,5,"ruler",2],[1,5,"caregiver",2],[2,1,"explorer",2],[2,1,"outlaw",1],[2,2,"explorer",1],[2,2,"hero",1],[2,4,"innocent",1],[2,4,"caregiver",1],[2,5,"innocent",2],[2,5,"caregiver",2],[3,1,"caregiver",2],[3,1,"lover",1],[3,2,"caregiver",1],[3,2,"regular",1],[3,4,"hero",1],[3,4,"ruler",1],[3,5,"hero",2],[3,5,"ruler",1],[4,1,"hero",2],[4,1,"magician",1],[4,1,"sage",1],[4,2,"hero",1],[4,2,"magician",1],[4,4,"regular",1],[4,4,"caregiver",1],[4,5,"regular",2],[4,5,"caregiver",2],[5,1,"regular",2],[5,1,"explorer",1],[5,1,"innocent",1],[5,2,"regular",1],[5,2,"explorer",1],[5,4,"ruler",1],[5,4,"hero",1],[5,5,"ruler",2],[5,5,"hero",1],[6,1,"creator",2],[6,1,"magician",2],[6,2,"creator",1],[6,2,"magician",1],[6,4,"sage",1],[6,4,"ruler",1],[6,5,"sage",2],[6,5,"ruler",1],[7,1,"outlaw",2],[7,1,"magician",1],[7,1,"hero",1],[7,2,"outlaw",1],[7,2,"magician",1],[7,4,"innocent",1],[7,4,"ruler",1],[7,4,"caregiver",1],[7,5,"innocent",2],[7,5,"ruler",1],[7,5,"caregiver",1],[8,1,"caregiver",2],[8,1,"lover",1],[8,1,"regular",1],[8,2,"caregiver",1],[8,2,"regular",1],[8,4,"sage",1],[8,4,"ruler",1],[8,5,"sage",2],[8,5,"ruler",1],[9,1,"regular",2],[9,1,"caregiver",2],[9,2,"regular",1],[9,2,"caregiver",1],[9,4,"hero",1],[9,4,"ruler",1],[9,5,"hero",2],[9,5,"ruler",2],[10,1,"magician",2],[10,1,"hero",1],[10,1,"sage",1],[10,2,"magician",1],[10,2,"sage",1],[10,4,"caregiver",1],[10,4,"innocent",1],[10,5,"caregiver",2],[10,5,"innocent",2],[11,1,"magician",2],[11,1,"creator",1],[11,1,"lover",1],[11,2,"magician",1],[11,2,"creator",1],[11,4,"sage",1],[11,4,"regular",1],[11,5,"sage",2],[11,5,"regular",2],[12,1,"explorer",2],[12,1,"outlaw",1],[12,2,"explorer",1],[12,2,"hero",1],[12,4,"caregiver",1],[12,4,"innocent",1],[12,5,"caregiver",2],[12,5,"innocent",2],[13,1,"regular",2],[13,1,"jester",1],[13,2,"regular",1],[13,2,"caregiver",1],[13,4,"sage",1],[13,4,"ruler",1],[13,5,"sage",2],[13,5,"ruler",1],[14,1,"explorer",2],[14,1,"hero",1],[14,2,"explorer",1],[14,2,"creator",1],[14,4,"ruler",1],[14,4,"innocent",1],[14,5,"ruler",2],[14,5,"sage",1],[15,1,"creator",2],[15,1,"outlaw",1],[15,1,"magician",1],[15,2,"creator",1],[15,2,"outlaw",1],[15,4,"regular",1],[15,4,"innocent",1],[15,4,"caregiver",1],[15,5,"regular",2],[15,5,"innocent",2],[15,5,"caregiver",1],[16,1,"ruler",2],[16,1,"hero",1],[16,1,"sage",1],[16,2,"ruler",1],[16,2,"hero",1],[16,4,"regular",1],[16,4,"explorer",1],[16,5,"regular",2],[16,5,"explorer",2],[17,1,"ruler",2],[17,1,"sage",1],[17,2,"ruler",1],[17,2,"hero",1],[17,4,"lover",1],[17,4,"outlaw",1],[17,4,"jester",1],[17,5,"lover",2],[17,5,"outlaw",1],[17,5,"jester",1],[18,1,"hero",2],[18,1,"magician",1],[18,1,"sage",1],[18,2,"hero",1],[18,2,"magician",1],[18,4,"caregiver",1],[18,4,"regular",1],[18,5,"caregiver",2],[18,5,"regular",2],[19,1,"jester",2],[19,1,"innocent",1],[19,2,"jester",1],[19,2,"lover",1],[19,4,"sage",1],[19,4,"ruler",1],[19,5,"sage",2],[19,5,"ruler",1],[20,1,"outlaw",2],[20,1,"creator",1],[20,2,"outlaw",1],[20,2,"hero",1],[20,4,"innocent",1],[20,4,"ruler",1],[20,4,"caregiver",1],[20,5,"innocent",2],[20,5,"ruler",1],[20,5,"caregiver",1],[21,1,"sage",2],[21,1,"hero",1],[21,1,"magician",1],[21,2,"sage",1],[21,2,"hero",1],[21,4,"regular",1],[21,4,"caregiver",1],[21,5,"regular",2],[21,5,"caregiver",2],[22,1,"magician",2],[22,1,"hero",1],[22,1,"creator",1],[22,2,"magician",1],[22,2,"hero",1],[22,4,"regular",1],[22,4,"sage",1],[22,5,"regular",2],[22,5,"sage",1],[23,1,"explorer",2],[23,1,"hero",1],[23,1,"magician",1],[23,2,"explorer",1],[23,2,"hero",1],[23,4,"innocent",1],[23,4,"caregiver",1],[23,5,"innocent",2],[23,5,"caregiver",2],[24,1,"jester",2],[24,1,"lover",1],[24,2,"jester",1],[24,2,"regular",1],[24,4,"sage",1],[24,4,"ruler",1],[24,4,"hero",1],[24,5,"sage",2],[24,5,"ruler",1],[24,5,"hero",1],[25,1,"outlaw",2],[25,1,"hero",1],[25,2,"outlaw",1],[25,2,"hero",1],[25,4,"ruler",1],[25,4,"sage",1],[25,5,"ruler",2],[25,5,"sage",1],[26,1,"lover",2],[26,1,"jester",1],[26,1,"caregiver",1],[26,2,"lover",1],[26,2,"caregiver",1],[26,4,"sage",1],[26,4,"ruler",1],[26,5,"sage",2],[26,5,"ruler",2],[27,1,"lover",2],[27,1,"magician",1],[27,2,"lover",1],[27,2,"creator",1],[27,4,"regular",1],[27,4,"sage",1],[27,5,"regular",2],[27,5,"sage",1],[28,1,"magician",2],[28,1,"jester",1],[28,1,"outlaw",1],[28,2,"magician",1],[28,2,"creator",1],[28,4,"caregiver",1],[28,4,"innocent",1],[28,5,"caregiver",2],[28,5,"innocent",2],[29,1,"creator",2],[29,1,"magician",1],[29,1,"lover",1],[29,2,"creator",1],[29,2,"magician",1],[29,4,"sage",1],[29,4,"ruler",1],[29,5,"sage",2],[29,5,"ruler",2],[30,1,"creator",2],[30,1,"outlaw",1],[30,1,"explorer",1],[30,2,"creator",1],[30,2,"outlaw",1],[30,4,"ruler",1],[30,4,"sage",1],[30,5,"ruler",2],[30,5,"sage",1]]';
const COMBINATIONS_DATA = [
  [`primary`, `secondary`, `type`, `description`],
  [`innocent`, `sage`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Wijze als secundair archetype is iemand die in het goede gelooft en die overtuiging onderbouwt met kennis. Het optimisme staat voorop, maar het is geen naïef optimisme. De Wijze brengt de feiten, het onderzoek, de zorgvuldigheid die het vertrouwen rechtvaardigt. Denk aan een merk dat eerlijk en hoopvol is, en tegelijk laat zien dat het zijn huiswerk heeft gedaan.
 De toon is positief en helder, met inhoudelijke onderbouwing waar het ertoe doet. Beelden zijn licht en clean. Boodschappen zijn rechtdoorzee, maar verwijzen naar bronnen, onderzoek of expertise wanneer geloofwaardigheid telt. Het merk legt uit, maar zonder belerend te worden. Eenvoud blijft leidend, kennis is dienstbaar.
 De combinatie is krachtig omdat ze warmte verbindt met geloofwaardigheid. Het merk komt over als oprecht én betrouwbaar, een zeldzame combinatie. De valkuil is dat de Wijze de Onschuldige overschaduwt en de communicatie te educatief wordt, waardoor de eenvoud en het vertrouwen verloren gaan. Of andersom: dat de Onschuldige de Wijze afzwakt tot oppervlakkige geruststelling zonder echte inhoud.`],
  [`innocent`, `explorer`, `productieve spanning`, `Een merk met de Onschuldige als kern en de Ontdekker als secundair archetype is iemand die in puurheid gelooft en die puurheid juist buiten de gebaande paden zoekt. Het is hoopvol én avontuurlijk, geruststellend én uitnodigend. De Onschuldige zorgt voor veiligheid en eerlijkheid, de Ontdekker brengt nieuwsgierigheid en onafhankelijkheid. Denk aan een merk dat zegt: ga op pad, blijf wie je bent.
 De toon is licht en open, met een uitnodiging om verder te kijken. Beelden tonen ruimte, natuur en eenvoudige ervaringen. Boodschappen gaan over puur plezier, eerlijke ontdekkingen en authentieke momenten. Het merk romantiseert niet, maar laat zien dat het simpele en het avontuurlijke samengaan.
 De combinatie is krachtig omdat ze verlangen koppelt aan vertrouwen. Het merk biedt vrijheid zonder risico, ontdekking zonder cynisme. De valkuil is dat de twee archetypen elkaar verzwakken: de Onschuldige maakt de Ontdekker tam, de Ontdekker maakt de Onschuldige zweverig. Het resultaat kan dan vrijblijvend worden, een merk dat noch echt geruststelt noch echt prikkelt.`],
  [`innocent`, `outlaw`, `niet eenvoudig`, `Een merk met de Onschuldige als kern en de Rebel als secundair archetype is iemand die in eerlijkheid gelooft en bereid is daarvoor de regels te breken. Het is een zachte rebel, of een principiële optimist. De Onschuldige geeft het morele kompas, de Rebel geeft de moed om tegen de stroom in te gaan wanneer de wereld niet eerlijk is. Tony's Chocolonely is hiervan het schoolvoorbeeld.
 De toon is helder en menselijk, met scherpte waar het over onrecht gaat. Beelden zijn vaak warm en herkenbaar, maar de boodschap kan confronteren. Het merk benoemt misstanden, maar vanuit hoop, niet vanuit cynisme. Humor is mogelijk, sarcasme niet. De rebellie is dienstbaar aan een puur ideaal.
 De combinatie is krachtig omdat ze morele autoriteit verbindt met daadkracht. Het merk durft positie te kiezen zonder zijn warmte te verliezen. De valkuil zit in de balans: te veel Rebel maakt het merk hard en verliest het vertrouwen, te veel Onschuldige maakt de rebellie krachteloos en moralistisch. Deze combinatie vraagt voortdurende afstemming tussen overtuigingskracht en zachtheid.`],
  [`innocent`, `magician`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Magiër als secundair archetype is iemand die in het goede gelooft en die geloof transformeert in iets bijzonders. Het is hoopvol én betoverend, geruststellend én verwonderlijk. De Onschuldige geeft de zuiverheid, de Magiër maakt er iets magisch van. Disney is een natuurlijke vertegenwoordiger van deze combinatie.
 De toon is warm en wonderlijk. Beelden zijn licht, sprookjesachtig of bijna kinderlijk verwonderd. Boodschappen gaan over dromen die uitkomen, over schoonheid in eenvoud, over de magie van het alledaagse. Het merk vertelt verhalen, geen verkooppraatjes. De belofte is emotioneel: hier mag je geloven.
 De combinatie is krachtig omdat ze pure emotie levert. Het merk raakt mensen op een diep gevoelsniveau. De valkuil is dat het te zoet wordt, te ver verwijderd van de werkelijkheid, of dat het claims doet die het niet kan waarmaken. Magie zonder geloofwaardigheid wordt kitsch, onschuld zonder anker wordt naïef.`],
  [`innocent`, `hero`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Held als secundair archetype is iemand die gelooft dat het goede uiteindelijk wint, mits je ervoor strijdt. De Onschuldige geeft de zuivere intentie, de Held de doorzettingskracht. Het is een merk dat met beide voeten op de grond staat en tegelijk aanmoedigt. Denk aan merken die optimisme koppelen aan prestatie.
 De toon is positief en motiverend. Beelden tonen mensen die iets bereiken, vaak in herkenbare en warme settings. Boodschappen gaan over volhouden, over het goede in mensen, over wat mogelijk is als je gelooft en doorzet. Het merk inspireert zonder te schreeuwen.
 De combinatie is krachtig omdat ze hoop koppelt aan actie. Het merk geeft mensen het gevoel dat ze het kunnen, en dat het de moeite waard is. De valkuil is dat het te braaf wordt, te voorspelbaar in zijn boodschap. Of dat de Held de Onschuldige overstemt, waardoor de communicatie meer over winnen gaat dan over geloven.`],
  [`innocent`, `lover`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Verleider als secundair archetype is iemand die in puurheid gelooft en die puurheid esthetisch viert. De Onschuldige geeft de eerlijkheid, de Verleider de schoonheid en intimiteit. Het is een merk dat warmte, genot en nabijheid biedt zonder pretentie. Denk aan natuurlijke verzorgingsmerken die zachtheid en zelfzorg combineren.
 De toon is warm, persoonlijk en uitnodigend. Beelden zijn esthetisch maar niet glamoureus, intiem maar niet gemaakt. Boodschappen gaan over kleine momenten van genot, over jezelf de tijd gunnen, over schoonheid in het eenvoudige. Het merk verleidt zachtjes, zonder druk.
 De combinatie is krachtig omdat ze emotionele warmte tastbaar maakt. Het merk voelt vertrouwd én verlokkend. De valkuil is dat de Verleider de Onschuldige geforceerd kan maken: te bewust mooi, te gepoetst, waardoor de eerlijkheid verloren gaat. Of dat de Onschuldige de Verleider afvlakt tot iets braafs zonder zinnelijke aantrekkingskracht.`],
  [`innocent`, `jester`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Joker als secundair archetype is iemand die het leven licht en eerlijk wil houden. De Onschuldige geeft het optimisme en vertrouwen, de Joker brengt humor en speelsheid. Het is een merk dat blij wordt van eenvoud en die blijdschap deelt zonder cynisme. Denk aan kindgerichte merken of merken met een lichte volwassen toon.
 De toon is vrolijk, warm en speels. Beelden zijn kleurrijk, vaak met een knipoog of onverwacht detail. Boodschappen zijn licht, soms grappig, maar nooit ten koste van iemand. Het merk lacht mét de doelgroep, niet om iets. Humor is vriendelijk en vindt aansluiting bij gedeelde herkenning.
 De combinatie is krachtig omdat ze plezier koppelt aan vertrouwen. Het merk maakt mensen blij op een manier die veilig voelt. De valkuil is dat het te kinderlijk wordt voor een volwassen doelgroep, of dat het te vriendelijk is om op te vallen. Zonder scherpte kan deze combinatie vervagen tot iets sympathieks zonder onderscheidend vermogen.`],
  [`innocent`, `regular`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Bondgenoot als secundair archetype is iemand die in eerlijkheid gelooft en in de waarde van gewoon zijn. De Onschuldige geeft de zuiverheid, de Bondgenoot de toegankelijkheid. Het is een merk zonder pretenties, dichtbij de gebruiker, oprecht en herkenbaar. Veel huismerken en familieverpakkingen leunen tegen deze combinatie aan.
 De toon is helder, vriendelijk en vertrouwd. Beelden tonen gewone mensen in alledaagse situaties. Boodschappen gaan over wat het leven mooi maakt: samen eten, kleine routines, eerlijke producten. Het merk spreekt geen reclametaal, maar gewone taal. Geruststelling zit in herkenning.
 De combinatie is krachtig omdat ze toegankelijkheid en betrouwbaarheid combineert. Het merk voelt als een goede buur. De valkuil is bleekheid: zonder spanning of profilering kan deze combinatie generiek worden, een merk dat overal en nergens is. De uitdaging is om in dat vriendelijke en eerlijke karakter toch een eigen stem te vinden.`],
  [`innocent`, `caregiver`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Beschermer als secundair archetype is iemand die in het goede gelooft en die overtuiging vertaalt naar zorg voor anderen. De Onschuldige geeft de hoopvolle blik, de Beschermer de toewijding om er voor mensen te zijn. Het is warm, betrouwbaar en oprecht behulpzaam. Veel zorgmerken en non-profits vallen in deze hoek.
 De toon is warm, geruststellend en empathisch. Beelden tonen menselijke nabijheid, zorg en eenvoudige momenten van verbinding. Boodschappen gaan over er zijn voor elkaar, over kleine gebaren met grote betekenis. Het merk legt geen druk op en moraliseert niet, maar nodigt uit tot zorgzaamheid.
 De combinatie is krachtig omdat ze emotionele warmte koppelt aan handelen. Het merk biedt hoop én concrete zorg. De valkuil is sentimentaliteit: te veel zachtheid kan de boodschap krachteloos maken, of het merk overkomen als naïef ten opzichte van de werkelijke complexiteit. De combinatie werkt het best wanneer de zorg concreet en de hoop verdiend is.`],
  [`innocent`, `ruler`, `niet eenvoudig`, `Een merk met de Onschuldige als kern en de Leider als secundair archetype is iemand die in eerlijkheid en eenvoud gelooft en tegelijk autoriteit en kwaliteit uitstraalt. De Onschuldige geeft het vertrouwen, de Leider geeft de standaard. Dit is geen voor de hand liggende combinatie, want eenvoud en autoriteit kunnen botsen. Maar bij merken met een lange traditie en een onbetwiste kwaliteit werkt het.
 De toon is helder, zelfverzekerd zonder arrogantie, en ingetogen krachtig. Beelden tonen kwaliteit, traditie en bewuste eenvoud. Boodschappen gaan over hoe het hoort, over wat blijft, over standaarden die ergens vandaan komen. Het merk is gezaghebbend zonder afstandelijk te worden.
 De combinatie is krachtig omdat ze geloofwaardigheid en aspiratie verbindt. Het merk wordt gezien als een betrouwbare standaard. De valkuil is dat de Leider de Onschuldige overstemt en de communicatie afstandelijk of elitair wordt. Of dat de Onschuldige de Leider verzwakt tot vrijblijvendheid. Deze combinatie vraagt zorgvuldige balans en een sterke merkgeschiedenis om geloofwaardig te blijven.`],
  [`innocent`, `creator`, `natuurlijke aanvulling`, `Een merk met de Onschuldige als kern en de Creator als secundair archetype is iemand die in puurheid gelooft en die puurheid uitdrukt in vorm en ambacht. De Onschuldige geeft de eerlijke intentie, de Creator de zorgvuldigheid en originaliteit. Het is een merk dat eenvoudig én verfijnd kan zijn, oprecht én esthetisch. Denk aan ambachtelijke merken die hun werk met liefde maken.
 De toon is warm en verzorgd. Beelden tonen ambacht, materiaal en eenvoudige schoonheid. Boodschappen gaan over hoe iets gemaakt is, over de mensen erachter, over de kleine keuzes die het verschil maken. Het merk laat zien dat zorgvuldigheid een vorm van eerlijkheid is.
 De combinatie is krachtig omdat ze authenticiteit en kwaliteit verbindt. Het merk straalt zorg en oprechtheid uit. De valkuil is preciositeit: te veel nadruk op vorm kan de eenvoud overschaduwen, of het merk laten overkomen als gewichtig. De combinatie werkt het best wanneer het ambacht dienstbaar blijft aan een eerlijk product.`],
  [`sage`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Onschuldige als secundair archetype is iemand die kennis deelt vanuit een geloof in het goede. De Wijze geeft de inhoud en autoriteit, de Onschuldige geeft de toegankelijkheid en oprechtheid. Het is een merk dat slim is zonder afstandelijk te zijn, deskundig zonder belerend. Denk aan informatiemerken met een menselijke toon.
 De toon is helder, onderbouwd en vriendelijk. Beelden zijn vaak professioneel maar warm, met aandacht voor mensen achter de feiten. Boodschappen leggen uit, maar zonder jargon waar het kan. Het merk educeert zonder te imponeren en blijft optimistisch over wat kennis kan brengen.
 De combinatie is krachtig omdat ze expertise koppelt aan toegankelijkheid. Het merk wordt vertrouwd én begrepen. De valkuil is dat de Wijze de Onschuldige overstemt en de communicatie te leerstellig wordt, of dat de Onschuldige de Wijze afvlakt tot oppervlakkigheid. De combinatie werkt het best wanneer kennis dienstbaar is aan helderheid en vertrouwen.`],
  [`sage`, `explorer`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Ontdekker als secundair archetype is iemand die kennis vergaart door te onderzoeken, te reizen, grenzen te verleggen. De Wijze geeft de analytische diepgang, de Ontdekker de nieuwsgierigheid en het lef. Het is een merk dat zoekt en vindt, en zijn vondsten deelt. Denk aan wetenschapsmerken, onderzoeksinstituten of vakbladen met een avontuurlijke ondertoon.
 De toon is gefundeerd én open, gezaghebbend én nieuwsgierig. Beelden tonen onderzoek, expedities, ontdekkingen, mensen die op zoek gaan. Boodschappen gaan over wat we nog niet weten, over verder kijken, over de moed om vragen te stellen. Het merk daagt uit om verder te denken.
 De combinatie is krachtig omdat ze kennis levend maakt. Het merk staat voor doordachte avonturen en doordringende inzichten. De valkuil is dat de Ontdekker de Wijze laat verdwalen in losse fascinaties, of dat de Wijze de Ontdekker indamt tot academische droogheid. De combinatie werkt het best wanneer onderzoek echt iets oplevert.`],
  [`sage`, `outlaw`, `productieve spanning`, `Een merk met de Wijze als kern en de Rebel als secundair archetype is iemand die kennis gebruikt om de status quo uit te dagen. De Wijze geeft het denkraam en de onderbouwing, de Rebel de scherpte en de moed om aan heilige huisjes te schudden. Het is een merk dat met argumenten provoceert. Denk aan kritische denkers, onderzoeksjournalistiek of disruptieve experts.
 De toon is scherp en doordacht. Beelden zijn vaak strak en onorthodox tegelijk. Boodschappen gaan over wat anderen niet durven te zeggen, over inzichten die ongemakkelijk zijn maar waar. Het merk argumenteert hard, maar onderbouwt elk standpunt. Provocatie is gefundeerd, niet gratuit.
 De combinatie is krachtig omdat ze autoriteit verbindt aan moed. Het merk wordt zowel gerespecteerd als gevreesd. De valkuil is overmoed: kennis kan overgaan in gelijkhebberij, rebellie kan de inhoud overschaduwen. De combinatie werkt het best wanneer er werkelijk iets te zeggen is en de provocatie iets oplost dat zonder moed niet wordt aangepakt.`],
  [`sage`, `magician`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Magiër als secundair archetype is iemand die kennis inzet om transformatie mogelijk te maken. De Wijze geeft de wetenschap en het inzicht, de Magiër de visie en de belofte van verandering. Het is een merk dat begrijpt en omvormt. Denk aan technologie- of innovatiemerken die kennis inzetten voor doorbraken.
 De toon is doordacht en visionair. Beelden zijn strak, vaak met een toekomstgerichte esthetiek. Boodschappen gaan over wat mogelijk wordt als we beter begrijpen, over hoe inzicht de wereld verandert. Het merk legt uit én belooft, met de inhoud als bewijs voor de visie.
 De combinatie is krachtig omdat ze geloofwaardigheid koppelt aan ambitie. Het merk biedt onderbouwde wonderen. De valkuil is dat de Magiër de Wijze laat overdrijven in beloftes die de feiten niet aankunnen, of dat de Wijze de Magiër terugbrengt tot voorzichtigheid. De combinatie werkt het best wanneer de transformatie echt door kennis wordt gedragen.`],
  [`sage`, `hero`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Held als secundair archetype is iemand die kennis inzet om te presteren en te overwinnen. De Wijze geeft de strategie, de Held de uitvoering. Het is een merk dat wint omdat het beter doordacht is. Denk aan sportwetenschappelijke merken, prestatiegedreven adviesbureaus of trainingsmethoden die op data steunen.
 De toon is gefocust en zelfverzekerd. Beelden tonen prestatie en analyse, vaak naast elkaar. Boodschappen gaan over wat je kunt bereiken als je het juiste weet en doet. Het merk koppelt inzicht aan resultaat. Geen retoriek zonder bewijs, geen analyse zonder consequentie.
 De combinatie is krachtig omdat ze denken en doen verbindt. Het merk wint met argumenten en met daadkracht. De valkuil is hardheid: de combinatie kan kil overkomen, alleen gericht op resultaat. Of de Wijze remt de Held met te veel nuance. De combinatie werkt het best wanneer kennis en prestatie elkaar voortdurend voeden.`],
  [`sage`, `lover`, `niet eenvoudig`, `Een merk met de Wijze als kern en de Verleider als secundair archetype is iemand die kennis presenteert met smaak en verfijning. De Wijze geeft de inhoud, de Verleider de stijl en de zinnelijkheid. Het is een merk dat denken aantrekkelijk maakt, een merk waarin esthetiek en inzicht samengaan. Denk aan culturele merken, premium magazines of merken die intellectueel genot bieden.
 De toon is verzorgd, intelligent en aantrekkelijk. Beelden zijn esthetisch én betekenisvol. Boodschappen gaan over schoonheid in inzicht, over de plek van smaak in denken. Het merk verleidt door zijn intelligentie en omgekeerd. Het wil dat lezen, kijken of luisteren een genoegen is.
 De combinatie is krachtig omdat ze de zintuigen en het denken samenbrengt. Het merk maakt iets aantrekkelijk wat anders droog zou zijn. De valkuil is gemaaktheid: te veel stijl maakt de inhoud verdacht, te veel inhoud maakt de stijl pretentieus. De combinatie vraagt voortdurende afstemming tussen de hoofd- en zijlijnen om geloofwaardig te blijven.`],
  [`sage`, `jester`, `productieve spanning`, `Een merk met de Wijze als kern en de Joker als secundair archetype is iemand die kennis met humor verpakt. De Wijze geeft het inzicht, de Joker maakt het toegankelijk en lichtvoetig. Het is een merk dat slim genoeg is om zichzelf niet te serieus te nemen. Denk aan slimme uitlegmerken, kennisplatformen met een eigen stem of educatieve merken die durven lachen.
 De toon is geestig én scherp. Beelden zijn vaak grafisch helder met een knipoog. Boodschappen gaan over kennis die niet zwaar hoeft te zijn, over inzichten die je kunt delen op een feestje. Het merk gebruikt humor om iets te onthullen, niet om iets te verbloemen. De grap bevestigt de inhoud.
 De combinatie is krachtig omdat ze onthouden wordt. Een grap die iets uitlegt, blijft hangen. De valkuil is dat de Joker de Wijze ondermijnt: als alles een grap wordt, neemt niemand de inhoud nog serieus. Of de Wijze maakt de Joker stijf. De combinatie werkt het best wanneer humor en helderheid hetzelfde doel dienen.`],
  [`sage`, `regular`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Bondgenoot als secundair archetype is iemand die kennis deelt op ooghoogte. De Wijze geeft de expertise, de Bondgenoot de toegankelijkheid en herkenbare taal. Het is een merk dat slimme dingen op een gewone manier zegt. Denk aan vakbladen met een brede lezerskring, of educatieve platformen die ingewikkelde dingen begrijpelijk maken.
 De toon is helder, vriendelijk en zonder pretentie. Beelden tonen mensen, voorbeelden, situaties uit het dagelijks leven. Boodschappen leggen uit met praktische voorbeelden en herkenbare taal. Het merk vermijdt jargon, maar laat de complexiteit niet wegvallen. Begrijpelijk zonder simplistisch te worden.
 De combinatie is krachtig omdat ze kennis democratiseert. Het merk geeft mensen het gevoel dat ze het kunnen begrijpen. De valkuil is dat de inhoud te veel wordt afgevlakt, of dat het merk te veel naar het midden trekt en zijn scherpte verliest. De combinatie werkt het best wanneer toegankelijkheid niet ten koste gaat van inhoudelijke diepgang.`],
  [`sage`, `caregiver`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Beschermer als secundair archetype is iemand die kennis inzet om mensen te helpen en te behoeden. De Wijze geeft de expertise, de Beschermer de zorg en de empathie. Het is een merk dat begrijpt wat er aan de hand is en handelt vanuit verantwoordelijkheid. Denk aan medische merken, juridische adviesbureaus of zorgverzekeraars met een sterke inhoudelijke poot.
 De toon is rustig, deskundig en empathisch. Beelden tonen aandacht, deskundigheid en menselijke verbondenheid. Boodschappen gaan over begrijpen, voorbereiden en bijstaan. Het merk legt uit waar het ingewikkeld is en stelt gerust waar het kan. Inhoud en zorg zijn één.
 De combinatie is krachtig omdat ze deskundigheid menselijk maakt. Het merk wordt vertrouwd op zowel inhoud als intentie. De valkuil is paternalisme: de Wijze kan belerend worden, de Beschermer betuttelend. De combinatie werkt het best wanneer kennis aanbiedt zonder op te leggen, en zorg ondersteunt zonder over te nemen.`],
  [`sage`, `ruler`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Leider als secundair archetype is iemand die kennis omzet in autoriteit en standaarden. De Wijze geeft de expertise, de Leider de stelligheid en de positie. Het is een merk dat de markt vormt door te weten en te bepalen. Denk aan toonaangevende adviesbureaus, normerende instituten of high-end vakmerken.
 De toon is gezaghebbend, doordacht en zelfbewust. Beelden zijn strak, vaak in gedempte of klassieke esthetiek. Boodschappen gaan over wat de norm is, wat de standaard zou moeten zijn, wat we van een sector mogen verwachten. Het merk spreekt zonder te overtuigen, omdat het zijn positie heeft verdiend.
 De combinatie is krachtig omdat ze gezag onderbouwt. Het merk leidt op basis van inhoud, niet alleen op basis van positie. De valkuil is afstandelijkheid en zelfgenoegzaamheid: de combinatie kan zo serieus worden dat ze het contact met de werkelijkheid verliest. Of zo zelfverzekerd dat ze niet meer leert. De combinatie werkt het best wanneer het leiderschap nog altijd ergens vandaan komt.`],
  [`sage`, `creator`, `natuurlijke aanvulling`, `Een merk met de Wijze als kern en de Creator als secundair archetype is iemand die kennis vertaalt in vorm en uitvinding. De Wijze geeft het inzicht, de Creator de toepassing en het ontwerp. Het is een merk dat denkt en maakt. Denk aan researchgedreven designbureaus, technologische pioniers of academische makers.
 De toon is intelligent en gemaakt met aandacht. Beelden zijn vaak verzorgd, met een grafische of conceptuele kant. Boodschappen gaan over hoe een idee vorm krijgt, over de relatie tussen denken en maken. Het merk laat zijn werk zien als bewijs van zijn denken.
 De combinatie is krachtig omdat ze conceptueel werk geloofwaardig maakt. Het merk laat zien dat de vorm volgt uit het denken. De valkuil is dat de Creator de Wijze afleidt tot esthetisch spel zonder inhoud, of dat de Wijze de Creator beperkt tot illustratie van inzichten. De combinatie werkt het best wanneer denken en maken hetzelfde proces zijn.`],
  [`explorer`, `innocent`, `productieve spanning`, `Een merk met de Ontdekker als kern en de Onschuldige als secundair archetype is iemand die de wereld in trekt met een open blik. De Ontdekker geeft de drang naar nieuwe ervaringen, de Onschuldige geeft de zuiverheid en het vertrouwen. Het is een merk dat avontuur biedt zonder cynisme. Denk aan natuurmerken, ecotoerisme of merken die ontdekken zonder veroveren.
 De toon is uitnodigend, helder en hoopvol. Beelden tonen weidsheid, natuur, verwondering. Boodschappen gaan over zelf ontdekken, over open staan voor wat komt, over de eenvoud van het pad. Het merk romantiseert niet, maar straalt vertrouwen in de wereld uit.
 De combinatie is krachtig omdat ze avontuur toegankelijk maakt. Het merk geeft mensen het gevoel dat ze op pad kunnen, ook als ze geen avonturier zijn. De valkuil is dat de Onschuldige de Ontdekker afvlakt tot oppervlakkige escapism, of dat de Ontdekker de Onschuldige overneemt en het merk te wild wordt voor zijn doelgroep. De combinatie werkt het best wanneer ontdekken een stille kracht is.`],
  [`explorer`, `sage`, `natuurlijke aanvulling`, `Een merk met de Ontdekker als kern en de Wijze als secundair archetype is iemand die op pad gaat om te leren. De Ontdekker geeft de beweging en de openheid, de Wijze geeft de reflectie en de duiding. Het is een merk dat avontuur en inzicht verbindt. Denk aan reismagazines met diepgang, antropologische merken of expeditie-organisaties.
 De toon is nieuwsgierig en gefundeerd. Beelden tonen onderweg-zijn én observatie, plekken én betekenis. Boodschappen gaan over wat een reis je leert, over de verhalen achter de plekken, over hoe ervaringen kennis worden. Het merk vertelt en denkt na.
 De combinatie is krachtig omdat ze diepte geeft aan beleving. Het merk biedt meer dan een ervaring, het biedt inzicht. De valkuil is dat de Wijze de Ontdekker te zwaar maakt, of dat de Ontdekker oppervlakkig blijft wanneer de Wijze niet doorzet. De combinatie werkt het best wanneer ontdekken en begrijpen elkaars motor zijn.`],
  [`explorer`, `outlaw`, `natuurlijke aanvulling`, `Een merk met de Ontdekker als kern en de Rebel als secundair archetype is iemand die zijn eigen pad zoekt, tegen de stroom in. De Ontdekker geeft de drang naar het ongebaande, de Rebel de scherpte om te breken met wat moet. Het is een merk voor mensen die niet meegaan in de norm. Denk aan motormerken, alternatieve reismerken of niche-outdoorlabels.
 De toon is direct, vrij en ongepolijst. Beelden tonen ruige landschappen, eigenwijze mensen, plekken zonder toeristen. Boodschappen gaan over je eigen route kiezen, over weigeren mee te gaan in georganiseerde ervaringen. Het merk spreekt mensen aan die liever vies dan gepolijst zijn.
 De combinatie is krachtig omdat ze authentieke vrijheid uitdraagt. Het merk staat ergens voor en is niet bang het te tonen. De valkuil is exclusiviteit en pose: de combinatie kan een houding worden zonder substantie, of zo eigenwijs dat ze niemand meer aanspreekt. De combinatie werkt het best wanneer rebellie iets oplevert, niet alleen iets uitstraalt.`],
  [`explorer`, `magician`, `natuurlijke aanvulling`, `Een merk met de Ontdekker als kern en de Magiër als secundair archetype is iemand die op zoek gaat naar transformerende ervaringen. De Ontdekker geeft de beweging, de Magiër de belofte van verandering. Het is een merk dat ervaringen biedt die mensen anders maken. Denk aan retraite-organisaties, transformatieve reizen of avontuurlijke wellness.
 De toon is uitnodigend en betekenisvol. Beelden tonen mensen op een drempel, plekken met lading, momenten van verschuiving. Boodschappen gaan over wat een ervaring met je doet, over wat je meeneemt naar huis. Het merk belooft geen verandering, maar nodigt uit tot openheid.
 De combinatie is krachtig omdat ze ervaringen lading geeft. Het merk verkoopt geen reis, maar een verschuiving. De valkuil is overdrijving: de Magiër kan de Ontdekker tot pretentieuze beloften verleiden, of de Ontdekker maakt de Magiër vluchtig en oppervlakkig. De combinatie werkt het best wanneer de belofte van transformatie waargemaakt wordt door echte ervaring.`],
  [`explorer`, `hero`, `natuurlijke aanvulling`, `Een merk met de Ontdekker als kern en de Held als secundair archetype is iemand die op pad gaat om grenzen te verleggen. De Ontdekker geeft de drang, de Held het doorzettingsvermogen en de prestatie. Het is een merk voor wie verder wil gaan, hoger, dieper, langer. Denk aan outdoor-prestatiemerken, expeditiemerken of avontuurlijke sportlabels.
 De toon is gedreven en helder. Beelden tonen mensen op uitdagende plekken, in actie, soms uitgeput maar voldaan. Boodschappen gaan over wat mogelijk is als je doorzet, over de waarde van inspanning. Het merk inspireert door te tonen wat anderen hebben gedaan.
 De combinatie is krachtig omdat ze ambitie en avontuur verbindt. Het merk geeft mensen een reden om hun comfortzone te verlaten. De valkuil is dat de Held de Ontdekker overneemt en het alleen nog over presteren gaat, niet meer over verwondering. Of dat de Ontdekker de Held vervaagt tot vrijblijvend rondzwerven. De combinatie werkt het best wanneer de inspanning beloond wordt met betekenis.`],
  [`explorer`, `lover`, `productieve spanning`, `Een merk met de Ontdekker als kern en de Verleider als secundair archetype is iemand die schoonheid zoekt in het ongebaande. De Ontdekker geeft de openheid en de zoektocht, de Verleider de esthetiek en zinnelijkheid. Het is een merk dat avonturen presenteert als zinnelijke ervaringen. Denk aan luxe reizen, sensorisch design op locatie of merken die ruwe natuur en verfijning combineren.
 De toon is sensorisch en uitnodigend. Beelden zijn esthetisch én rauw, mooi én onverwacht. Boodschappen gaan over de schoonheid van het ongepolijste, over genot in ontdekken, over wat je voelt in plaats van alleen wat je ziet. Het merk laat ervaring spreken.
 De combinatie is krachtig omdat ze ontdekken aantrekkelijk maakt voor mensen die schoonheid zoeken. De valkuil is dat de Verleider de Ontdekker tam maakt, dat avontuur een esthetische pose wordt zonder echte beweging. Of dat de Ontdekker de Verleider verwart met verfijning. De combinatie werkt het best wanneer de zinnelijkheid uit de plek zelf komt, niet uit de presentatie ervan.`],
  [`explorer`, `jester`, `natuurlijke aanvulling`, `Een merk met de Ontdekker als kern en de Joker als secundair archetype is iemand die met plezier de wereld in trekt. De Ontdekker geeft de openheid, de Joker de speelsheid en de relativering. Het is een merk dat avontuur licht maakt, zonder het te bagatelliseren. Denk aan reisplatformen voor jongeren, festivalmerken of avontuurlijke voedselconcepten.
 De toon is opgewekt en ongedwongen. Beelden tonen mensen die plezier hebben onderweg, met scheve glimlachen en onverwachte momenten. Boodschappen gaan over avontuur dat geen heldenreis hoeft te zijn, over plezier als motivatie, over niet-zo-bedoelde ontdekkingen. Het merk lacht onderweg.
 De combinatie is krachtig omdat ze ontdekken laagdrempelig en aantrekkelijk maakt. Het merk neemt zichzelf niet te serieus en nodigt mensen uit dat ook niet te doen. De valkuil is dat het te oppervlakkig wordt, of dat de Joker de Ontdekker afzwakt tot consumentisme. De combinatie werkt het best wanneer de speelsheid uit echte beleving komt.`],
  [`explorer`, `regular`, `natuurlijke aanvulling`, `Een merk met de Ontdekker als kern en de Bondgenoot als secundair archetype is iemand die op pad gaat met gewone mensen, voor gewone mensen. De Ontdekker geeft de horizon, de Bondgenoot de toegankelijkheid en herkenbaarheid. Het is een merk dat ontdekking dichtbij brengt. Denk aan reisorganisaties zonder pretentie of outdoor-merken voor het gezin.
 De toon is open, vriendelijk en zonder afstand. Beelden tonen herkenbare mensen onderweg, niet alleen avonturiers. Boodschappen gaan over wat ontdekken kan zijn voor wie het nog nooit heeft gedaan, over kleine reizen met grote betekenis. Het merk is gids zonder leraar te zijn.
 De combinatie is krachtig omdat ze ontdekking democratiseert. Het merk geeft een breed publiek toegang tot wat avontuur lijkt. De valkuil is bleekheid: de combinatie kan zo herkenbaar worden dat ze haar ontdekkingsdrang verliest. Of zo gewoon dat ze niets meer onderscheidt. De combinatie werkt het best wanneer de toegankelijkheid niet de horizon verkleint.`],
  [`explorer`, `caregiver`, `productieve spanning`, `Een merk met de Ontdekker als kern en de Beschermer als secundair archetype is iemand die het ongebaande betreedt met zorg, voor mens, dier of plek. De Ontdekker geeft de drang naar buiten, de Beschermer de verantwoordelijkheid voor wat we tegenkomen. Het is een merk dat avontuur en zorg verbindt. Denk aan duurzame reisorganisaties, natuurbeschermingsprogramma's of buitenlessen met aandacht.
 De toon is bewust, betrokken en wijs. Beelden tonen plekken, mensen en dieren met aandacht voor hun kwetsbaarheid. Boodschappen gaan over hoe je iets ontdekt zonder het kapot te maken, over reizen met respect, over je rol in de wereld. Het merk denkt na.
 De combinatie is krachtig omdat ze beleving en moraal verbindt. Het merk laat zien dat avontuur niet ten koste hoeft te gaan van de wereld. De valkuil is moralisme: de Beschermer kan de Ontdekker schuldgevoel bezorgen en de ervaring doodslaan. Of de Ontdekker negeert de zorg en de combinatie wordt incoherent. De combinatie werkt het best wanneer de zorg natuurlijk in de ervaring is verweven.`],
  [`explorer`, `ruler`, `niet eenvoudig`, `Een merk met de Ontdekker als kern en de Leider als secundair archetype is iemand die zijn ontdekkingen tot standaard verheft. De Ontdekker geeft de pioniersgeest, de Leider de positie en het gezag. Het is een merk dat eerst ergens kwam en nu de norm bepaalt. Denk aan premium reismerken met erfgoed of vakgebieden waar ervaring autoriteit wordt.
 De toon is zelfverzekerd, ervaren en kalm. Beelden tonen exclusieve plekken en doorleefde mensen. Boodschappen gaan over wat échte ervaring is, over kwaliteit die je alleen krijgt door tijd, over standaarden die uit ontdekking zijn ontstaan. Het merk hoeft niet te overtuigen, het laat zijn track record spreken.
 De combinatie is krachtig omdat ze pioniersgeest tot premium verheft. De valkuil is dat de Leider de Ontdekker stilzet en het merk vastloopt in zijn verleden. Of dat de Ontdekker het gezag van de Leider ondermijnt door teveel beweging. De combinatie werkt het best wanneer de leider blijft ontdekken en zijn positie blijft verdienen, niet alleen claimt.`],
  [`explorer`, `creator`, `natuurlijke aanvulling`, `Een merk met de Ontdekker als kern en de Creator als secundair archetype is iemand die op pad gaat om iets nieuws te maken. De Ontdekker geeft de openheid, de Creator de drang om eruit iets te bouwen. Het is een merk dat ervaringen vertaalt in vorm of product. Denk aan ontwerpmerken die uit reizen geïnspireerd zijn, of foto- en filmstudio's met een sterke handtekening.
 De toon is observerend en gemaakt, met aandacht voor detail én voor het grotere geheel. Beelden tonen plekken, materialen, mensen die werken. Boodschappen gaan over wat een ervaring oplevert, over het ambacht van waarnemen en maken. Het merk laat zijn werk én zijn weg zien.
 De combinatie is krachtig omdat ze beweging en betekenis koppelt. Het merk maakt iets uit wat het tegenkomt. De valkuil is dat de Creator de Ontdekker laat verdwijnen achter het werk, of dat de Ontdekker de Creator afleidt van het maken. De combinatie werkt het best wanneer reis en werk dezelfde ritmiek hebben.`],
  [`outlaw`, `innocent`, `niet eenvoudig`, `Een merk met de Rebel als kern en de Onschuldige als secundair archetype is iemand die de regels breekt vanuit een puur ideaal. De Rebel geeft de scherpte en het lef, de Onschuldige het morele kompas. Het is een merk dat opstaat tegen onrecht zonder cynisch te worden. Patagonia en Tony's Chocolonely zitten in deze hoek, met de Rebel meer of minder dominant.
 De toon is direct, soms confronterend, maar warm in zijn kern. Beelden tonen misstanden én oprechte mensen. Boodschappen benoemen wat fout is, maar vanuit de overtuiging dat het anders kan. Het merk vecht voor iets, niet alleen tegen iets. Humor is mogelijk, maar gericht op het systeem, niet op individuen.
 De combinatie is krachtig omdat ze morele autoriteit verleent aan rebellie. De valkuil is moralisme: de Onschuldige kan de Rebel veranderen in een belerende activist, de Rebel kan de Onschuldige verbitteren. De combinatie werkt het best wanneer de strijd dienstbaar is aan een hoopvol ideaal en niet zelf doel wordt.`],
  [`outlaw`, `sage`, `productieve spanning`, `Een merk met de Rebel als kern en de Wijze als secundair archetype is iemand die met onderbouwing tegen de stroom in zwemt. De Rebel geeft de moed, de Wijze het argument. Het is een merk dat provoceert met inhoud. Denk aan kritische denkers, onderzoeksjournalistieke platforms of disruptieve consultants die met data komen.
 De toon is scherp, doordacht en goed gefundeerd. Beelden zijn vaak strak en gedurfd. Boodschappen gaan over wat anderen niet zien of niet durven te zeggen, met cijfers, voorbeelden en argumenten. Het merk valt aan, maar nooit zonder fundament.
 De combinatie is krachtig omdat ze rebellie geloofwaardig maakt. Het merk wordt serieus genomen, ook door wie het niet eens is. De valkuil is gelijkhebberij: de combinatie kan vervallen in eindeloze afrekening, of zo intellectueel worden dat de breedte van de doelgroep wegvalt. De combinatie werkt het best wanneer er werkelijk iets te zeggen is.`],
  [`outlaw`, `explorer`, `natuurlijke aanvulling`, `Een merk met de Rebel als kern en de Ontdekker als secundair archetype is iemand die zijn eigen weg gaat, ver van de gebaande paden. De Rebel geeft de breuk met het bestaande, de Ontdekker de drang naar het ongebaande. Het is een merk voor wie alleen wil reizen of bouwen aan iets nieuws. Denk aan motormerken, alternatieve labels of pioniers in een nichemarkt.
 De toon is ongepolijst, zelfverzekerd en eigenzinnig. Beelden tonen ruwe landschappen, eigenwijze mensen, plekken zonder publiek. Boodschappen gaan over je eigen pad kiezen, over weigeren mee te lopen, over het plezier van alleen zijn met je idee. Het merk is van de buitenstaander voor de buitenstaander.
 De combinatie is krachtig omdat ze authentieke afwijking verkoopt. Het merk staat ergens voor zonder te hoeven uitleggen waarom. De valkuil is pose: de combinatie kan een houding worden zonder werkelijke breuk, of zo eigenwijs dat ze niemand meer toelaat. De combinatie werkt het best wanneer de afwijking iets oplevert in plaats van alleen een statement te zijn.`],
  [`outlaw`, `magician`, `natuurlijke aanvulling`, `Een merk met de Rebel als kern en de Magiër als secundair archetype is iemand die het bestaande omverwerpt om iets nieuws mogelijk te maken. De Rebel geeft de breuk, de Magiër de visie van wat dan kan ontstaan. Het is een merk dat revolutie verbindt aan transformatie. Denk aan technologische pioniers die een industrie willen herzien.
 De toon is uitdagend en visionair. Beelden zijn strak, vaak met een toekomstgerichte esthetiek en een edge. Boodschappen gaan over wat we hebben aanvaard en niet meer hoeven te aanvaarden, over wat mogelijk wordt als we durven loslaten. Het merk verleidt met de ruimte voorbij de breuk.
 De combinatie is krachtig omdat ze verandering inspireert in plaats van afdwingt. De valkuil is grootspraak: de Magiër kan de Rebel verleiden tot beloften die niet kunnen, of de Rebel kan de Magiër laten lijken op een charlatan. De combinatie werkt het best wanneer de breuk verdiend wordt door wat erna komt.`],
  [`outlaw`, `hero`, `natuurlijke aanvulling`, `Een merk met de Rebel als kern en de Held als secundair archetype is iemand die strijdt tegen de gevestigde orde en wint. De Rebel geeft de breuk, de Held de overwinning en de prestatie. Het is een merk voor de underdog die de top haalt. Denk aan disruptieve sportmerken, of opstandige tech-pioniers die het opnemen tegen reuzen.
 De toon is strijdvaardig, energiek en doelgericht. Beelden tonen prestatie tegen de stroom in, mensen die het tegen de verwachting in halen. Boodschappen gaan over wat je kunt bereiken als je weigert mee te doen aan andermans regels. Het merk inspireert door zijn voorbeeld.
 De combinatie is krachtig omdat ze rebellie productief maakt. Het merk biedt mensen een kanaal voor hun ongenoegen, niet alleen een uitlaatklep. De valkuil is testosterongedreven retoriek: de combinatie kan vervallen in macho-taal die niet voor iedereen aansluit. Of de Rebel ondermijnt het succes van de Held door eeuwig in oppositie te blijven. De combinatie werkt het best wanneer winnen en breken in dezelfde beweging zitten.`],
  [`outlaw`, `lover`, `productieve spanning`, `Een merk met de Rebel als kern en de Verleider als secundair archetype is iemand die regels breekt met stijl en aantrekkingskracht. De Rebel geeft de tegendraadse kant, de Verleider de zinnelijkheid en het magnetisme. Het is een merk dat aantrekkelijk gevaarlijk is. Denk aan modemerken met een edgy signature, parfums met een rebelse claim of hospitality-merken die luxe en breuk combineren.
 De toon is scherp, zinnelijk en ongepolijst. Beelden zijn esthetisch én provocerend, mooi én ongemakkelijk. Boodschappen gaan over verlangen dat geen vergiffenis vraagt, over schoonheid voorbij de regels. Het merk verleidt door tegen de norm in te gaan.
 De combinatie is krachtig omdat ze begeerte en lef koppelt. Het merk wordt aantrekkelijk juist omdat het niet braaf is. De valkuil is provocatie zonder substantie: de combinatie kan oppervlakkig worden, een esthetische pose zonder werkelijke breuk. Of de Verleider maakt de Rebel commercieel tam. De combinatie werkt het best wanneer de aantrekkingskracht uit een echte breuk komt.`],
  [`outlaw`, `jester`, `natuurlijke aanvulling`, `Een merk met de Rebel als kern en de Joker als secundair archetype is iemand die het systeem belachelijk maakt om het te ontmaskeren. De Rebel geeft de tegendraadse kant, de Joker de humor en de ironie. Het is een merk dat lacht waar anderen zich druk maken, en daarmee scherpte legt op wat fout zit. Denk aan satirische platformen of merken die met humor de status quo aanvallen.
 De toon is geestig, brutaal en soms gemeen. Beelden zijn opvallend, vaak grafisch sterk met een twist. Boodschappen gaan over wat we niet hardop zeggen, over de absurditeit van de norm, over taboes met een lach. Het merk gebruikt humor als wapen, niet als versiering.
 De combinatie is krachtig omdat ze verzet ontwapenend maakt. Het merk wordt gedeeld omdat het lachen oproept en bij blijft. De valkuil is cynisme: de combinatie kan vervallen in spot zonder agenda, waardoor de rebellie hol wordt. Of de Joker zwakt de Rebel af tot vrijblijvende grollen. De combinatie werkt het best wanneer de grap iets blootlegt.`],
  [`outlaw`, `regular`, `productieve spanning`, `Een merk met de Rebel als kern en de Bondgenoot als secundair archetype is iemand die opstaat tegen het systeem namens de gewone mens. De Rebel geeft de strijd, de Bondgenoot de verbondenheid met wie het raakt. Het is een merk dat zijn rebellie niet voor zichzelf voert. Denk aan vakbondsachtige initiatieven, eerlijke prijzen-merken of activistische supermarkten.
 De toon is direct, herkenbaar en strijdbaar. Beelden tonen gewone mensen die hun stem verheffen, alledaagse situaties met een politieke laag. Boodschappen gaan over wat oneerlijk is voor mensen zoals jij en ik, over wat we samen kunnen veranderen. Het merk spreekt namens zijn doelgroep, niet boven hen.
 De combinatie is krachtig omdat ze rebellie verankert in solidariteit. Het merk wordt geloofd omdat het de pijn van zijn doelgroep deelt. De valkuil is populisme: de combinatie kan vervallen in oppervlakkige slogans, of zo nadrukkelijk over de gewone mens spreken dat ze hem niet meer hoort. De combinatie werkt het best wanneer de strijd echt voor mensen wordt gevoerd, niet over hen.`],
  [`outlaw`, `caregiver`, `productieve spanning`, `Een merk met de Rebel als kern en de Beschermer als secundair archetype is iemand die het opneemt voor wie kwetsbaar is. De Rebel geeft het verzet, de Beschermer de zorg en het beschermen. Het is een merk dat woedend wordt namens een ander. Denk aan dierenwelzijnsorganisaties, activistische zorginitiatieven of opstandige sociale ondernemers.
 De toon is fel én warm, scherp én betrokken. Beelden tonen kwetsbaarheid en strijd in dezelfde frame. Boodschappen gaan over wie geen stem heeft, over wat we niet meer accepteren, over wat we beschermen door op te staan. Het merk vecht uit zorg.
 De combinatie is krachtig omdat ze morele urgentie geeft aan rebellie. Het merk verzet zich vanuit een herkenbaar fundament. De valkuil is sentimentaliteit gekoppeld aan agressie: de combinatie kan zwaar of veroordelend worden. Of de Beschermer maakt de Rebel zacht en de strijd ineffectief. De combinatie werkt het best wanneer de bescherming concreet en de strijd gericht is.`],
  [`outlaw`, `ruler`, `productieve spanning`, `Een merk met de Rebel als kern en de Leider als secundair archetype is iemand die de markt heeft veranderd en nu zelf de standaard is. De Rebel geeft de oorsprong, de Leider de positie. Het is een merk dat ooit opstond tegen het bestaande en nu zelf het bestaande is, zonder zijn scherpte op te geven. Denk aan voormalige disruptors die marktleider zijn geworden.
 De toon is zelfverzekerd, scherp en met geheugen voor zijn afkomst. Beelden tonen succes met een edge, autoriteit met een stempel. Boodschappen gaan over hoe het anders werd, over wat het merk veranderd heeft, over de verantwoordelijkheid van succes. Het merk leidt zonder zich te conformeren.
 De combinatie is krachtig omdat ze rebellie volwassen maakt. Het merk laat zien dat verandering kan blijven. De valkuil is incoherentie: het merk kan worden ingehaald door de cultuur die het ooit zelf disrupteerde, of zo gevestigd raken dat de Rebel ongeloofwaardig wordt. De combinatie werkt het best wanneer het leiderschap gebruikt wordt om door te breken, niet om te conserveren.`],
  [`outlaw`, `creator`, `natuurlijke aanvulling`, `Een merk met de Rebel als kern en de Creator als secundair archetype is iemand die het bestaande afbreekt om iets nieuws te maken. De Rebel geeft de breuk, de Creator het bouwwerk dat erna komt. Het is een merk dat innovatie als vorm van verzet ziet. Denk aan disruptieve designers, alternatieve maakcollectieven of architecten met een politieke agenda.
 De toon is doortastend en doelgericht. Beelden tonen het werk én de breuk waaruit het ontstaat, processen én resultaten. Boodschappen gaan over waarom het bestaande niet voldoet en wat in plaats daarvan kan komen. Het merk maakt om te tonen wat anders mogelijk is.
 De combinatie is krachtig omdat ze constructief is. Het merk verzet zich met werk, niet alleen met woorden. De valkuil is dat de Rebel het maken vertraagt door alleen te willen breken, of dat de Creator de Rebel afzwakt tot esthetisch experiment zonder werkelijke breuk. De combinatie werkt het best wanneer afbreken en bouwen elkaars motor zijn.`],
  [`magician`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Onschuldige als secundair archetype is iemand die in transformatie gelooft vanuit een pure intentie. De Magiër geeft de visie, de Onschuldige de zuiverheid. Het is een merk dat magie biedt zonder cynisme. Denk aan kinderachtige verbeelding voor volwassenen, bewustzijnsmerken of speelgoed met diepgang.
 De toon is verwonderd, hoopvol en uitnodigend. Beelden zijn licht, vaak sprookjesachtig of conceptueel zacht. Boodschappen gaan over dromen die mogen bestaan, over magie in het alledaagse. Het merk vertelt verhalen die mensen weer doen geloven.
 De combinatie is krachtig omdat ze emotie en hoop stapelt. Het merk raakt diep zonder te manipuleren. De valkuil is sentimentaliteit: de combinatie kan te zoet worden, te veel beloven, of vervallen in fantasie zonder grond. Of de Onschuldige beperkt de Magiër tot iets veiligs. De combinatie werkt het best wanneer de magie iets echts onthult.`],
  [`magician`, `sage`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Wijze als secundair archetype is iemand die transformatie onderbouwt met inzicht. De Magiër geeft de visie, de Wijze de fundering. Het is een merk dat technologie of methode tot iets bijna magisch verheft. Denk aan high-tech innovatiemerken, of methodes waar wetenschap en verbeelding samengaan.
 De toon is gefundeerd én visionair. Beelden zijn strak en future-forward, soms met een geheimzinnige laag. Boodschappen gaan over wat mogelijk wordt door te begrijpen, over hoe inzicht de wereld verandert. Het merk legt uit én belooft, en houdt beide in balans.
 De combinatie is krachtig omdat ze ambitie geloofwaardig maakt. Het merk biedt magie met bewijs. De valkuil is overpromise: de Magiër kan de Wijze laten reiken naar wat het niet kan onderbouwen. Of de Wijze remt de Magiër en de visie verdwijnt. De combinatie werkt het best wanneer beloftes en kennis elkaar in dezelfde beweging dragen.`],
  [`magician`, `explorer`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Ontdekker als secundair archetype is iemand die transformatie zoekt door op pad te gaan. De Magiër geeft het verlangen naar verandering, de Ontdekker het zelf vinden ervan. Het is een merk dat groei aanbiedt via ervaring. Denk aan retraites, innerlijke reizen of leer- en groeitrajecten.
 De toon is uitnodigend, geheimzinnig en open. Beelden tonen drempels, plekken met lading, mensen op een keerpunt. Boodschappen gaan over wat je vindt als je gaat, over de moed om je leven anders te bekijken. Het merk roept op tot beweging, niet tot bekering.
 De combinatie is krachtig omdat ze transformatie persoonlijk maakt. Het merk geeft mensen het gevoel dat ze hun eigen verandering kunnen vinden. De valkuil is zweverigheid: de combinatie kan vervallen in vage spirituele beloftes. Of de Ontdekker maakt de Magiër vluchtig. De combinatie werkt het best wanneer de reis een echte verandering oplevert.`],
  [`magician`, `outlaw`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Rebel als secundair archetype is iemand die transformatie eist door het bestaande omver te werpen. De Magiër geeft de visie, de Rebel de breuk. Het is een merk dat verandering ziet als een doorbraak, niet als een verfijning. Denk aan disruptieve technologie, futuristische concepten of activistische innovators.
 De toon is uitdagend en visionair. Beelden zijn strak en futuristisch, vaak met een politieke laag. Boodschappen gaan over wat moet sneuvelen om iets nieuws te laten ontstaan, over de moed om het oude los te laten. Het merk verleidt met de wereld die voorbij de breuk ligt.
 De combinatie is krachtig omdat ze verandering urgent maakt. Het merk geeft mensen redenen om in te stappen op iets nieuws. De valkuil is overspanning: de combinatie kan vervallen in revolutionaire retoriek zonder onderbouwing, of beloftes doen die niet uitkomen. De combinatie werkt het best wanneer de breuk waargemaakt wordt door de visie erachter.`],
  [`magician`, `hero`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Held als secundair archetype is iemand die mensen helpt hun grootste versie te worden. De Magiër geeft de transformatie, de Held de prestatie en het doorzettingsvermogen. Het is een merk dat groei verbindt met overwinning. Denk aan coaching-merken, performance-platformen of merken die mensen anders maken door uitdaging.
 De toon is inspirerend, krachtig en geladen. Beelden tonen mensen die iets bereiken én iets worden. Boodschappen gaan over de versie van jezelf die je kunt worden, over wat aan de andere kant van inspanning ligt. Het merk wijst de weg én duwt op je rug.
 De combinatie is krachtig omdat ze ambitie en betekenis verbindt. Het merk biedt meer dan resultaat, het biedt verandering. De valkuil is grootspraak: de combinatie kan vervallen in motivational-speak. Of de Held verdringt de Magiër en het gaat alleen nog over presteren. De combinatie werkt het best wanneer de prestatie iemand echt verandert.`],
  [`magician`, `lover`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Verleider als secundair archetype is iemand die mensen toelaat te dromen door middel van schoonheid. De Magiër geeft de transformatie, de Verleider de zinnelijkheid en de aantrekkingskracht. Het is een merk dat verlangen koppelt aan verandering. Denk aan luxe parfummerken, premium hospitality of sensorische ervaringen die transformatie beloven.
 De toon is meeslepend, esthetisch en suggestief. Beelden zijn rijk, vaak met een mystieke kant. Boodschappen gaan over wie je wordt door te ervaren, over het wonder van de juiste plek of het juiste moment. Het merk bouwt een wereld waarin mensen willen verdwijnen.
 De combinatie is krachtig omdat ze verbeelding en verlangen koppelt. Het merk biedt geen product, het biedt een belofte van wie je kunt zijn. De valkuil is leegte: de combinatie kan vervallen in gepolijste fantasie zonder substantie. Of de Verleider maakt de Magiër zonder lading. De combinatie werkt het best wanneer de schoonheid daadwerkelijk verandering oplevert.`],
  [`magician`, `jester`, `productieve spanning`, `Een merk met de Magiër als kern en de Joker als secundair archetype is iemand die transformatie luchtig maakt. De Magiër geeft de visie, de Joker de relativering en de speelsheid. Het is een merk dat magie aanbiedt zonder pretenties. Denk aan creatieve platformen die innovatie speels brengen of merken die toekomst en humor combineren.
 De toon is verwonderd én lichtvoetig. Beelden zijn vaak verbeeldingsrijk en speels. Boodschappen gaan over toekomst die geen plechtige zaak hoeft te zijn, over verandering die ook leuk kan zijn. Het merk maakt grote ideeën klein toegankelijk.
 De combinatie is krachtig omdat ze transformatie aantrekkelijk maakt voor wie geschrokken is van zwaarte. De valkuil is dat de Joker de Magiër ondermijnt: als alles een grap is, blijft er geen visie over. Of de Magiër maakt de Joker plotseling te ernstig. De combinatie werkt het best wanneer de speelsheid de verbeelding versterkt.`],
  [`magician`, `regular`, `productieve spanning`, `Een merk met de Magiër als kern en de Bondgenoot als secundair archetype is iemand die transformatie binnen handbereik brengt voor de gewone mens. De Magiër geeft de belofte, de Bondgenoot de toegankelijkheid. Het is een merk dat zegt: ook jij. Denk aan democratische technologie, Apple in zijn vroege jaren, of platforms die professionele tools beschikbaar maken voor iedereen.
 De toon is uitnodigend en herkenbaar. Beelden tonen gewone mensen die iets buitengewoons doen, zonder afstand of pretenties. Boodschappen gaan over wat mogelijk wordt voor wie het tot voor kort niet kon, over de magie die niet alleen voor ingewijden is. Het merk slecht drempels.
 De combinatie is krachtig omdat ze ambitie democratiseert. Het merk geeft een breed publiek toegang tot wat exclusief leek. De valkuil is bleekheid: de Bondgenoot kan de magie wegvegen, de Magiër kan de Bondgenoot tot een marketing-truc maken. De combinatie werkt het best wanneer de toegankelijkheid niet ten koste gaat van de echte verandering die geboden wordt.`],
  [`magician`, `caregiver`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Beschermer als secundair archetype is iemand die transformatie aanbiedt vanuit zorg. De Magiër geeft de verandering, de Beschermer geeft de veiligheid waarbinnen die kan plaatsvinden. Het is een merk dat groei en zorg combineert. Denk aan therapeutische merken, mind-merken met aandacht of wellness met diepgang.
 De toon is rustig, warm en geheimzinnig in zachte zin. Beelden tonen mensen die in een veilige ruimte iets ondergaan, plekken met lading. Boodschappen gaan over wat je toelaat als je je veilig voelt, over de combinatie van overgave en zorg. Het merk biedt ruimte.
 De combinatie is krachtig omdat ze verandering veilig maakt. Het merk wordt vertrouwd voor wat anders eng zou zijn. De valkuil is dat de Beschermer de Magiër temt en de transformatie blijft uit, of dat de Magiër de Beschermer verleidt tot beloftes die de verantwoordelijkheid overstijgen. De combinatie werkt het best wanneer zorg en visie elkaar versterken.`],
  [`magician`, `ruler`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Leider als secundair archetype is iemand die de toekomst bepaalt vanuit zijn positie. De Magiër geeft de visie, de Leider de macht om die te realiseren. Het is een merk dat zegt: dit is waar de wereld heen gaat, en wij gaan voorop. Denk aan Tesla in zijn impact, of toonaangevende merken in opkomende sectoren.
 De toon is zelfverzekerd, visionair en gezaghebbend. Beelden tonen toekomst en exclusiviteit. Boodschappen gaan over wat morgen zal zijn, over standaard die nu wordt gezet, over leiderschap als bewijs van visie. Het merk laat geen ruimte voor twijfel.
 De combinatie is krachtig omdat ze ambitie vorm geeft. Het merk dwingt mensen serieus te nemen wat het zegt. De valkuil is hubris: de combinatie kan vervallen in zelfverklaarde profetie zonder grond. Of de Leider verdringt de Magiër en het wordt alleen nog macht. De combinatie werkt het best wanneer de visie en het leiderschap waargemaakt worden door consistente realisatie.`],
  [`magician`, `creator`, `natuurlijke aanvulling`, `Een merk met de Magiër als kern en de Creator als secundair archetype is iemand die zijn visie vorm geeft. De Magiër geeft de transformatieve verbeelding, de Creator het ambacht en de uitvoering. Het is een merk dat de toekomst niet alleen schetst maar ook bouwt. Denk aan visionaire designstudio's of pioniers in technologie die hun eigen wereld creëren.
 De toon is verbeeldingsrijk en doortastend. Beelden tonen werk dat een wereld oproept, processen die naar resultaat leiden. Boodschappen gaan over hoe een idee tot leven komt, over de relatie tussen denken en bouwen. Het merk laat zien wat het droomt door wat het maakt.
 De combinatie is krachtig omdat ze visie tastbaar maakt. Het merk levert wat het belooft. De valkuil is overdaad: de combinatie kan vervallen in artistieke pretentie. Of de Creator verdringt de Magiër en het werk wordt esthetisch maar zonder lading. De combinatie werkt het best wanneer de gemaakte dingen de verandering daadwerkelijk dragen.`],
  [`hero`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Onschuldige als secundair archetype is iemand die strijdt vanuit een eerlijk hart. De Held geeft de prestatie en de moed, de Onschuldige geeft de zuivere intentie. Het is een merk dat presteert zonder cynisme, dat overwint zonder op te scheppen. Denk aan sportmerken met een familiekarakter of merken die strijden voor het goede.
 De toon is positief, motiverend en oprecht. Beelden tonen mensen die zich inzetten, met aandacht voor herkenbare warmte. Boodschappen gaan over wat je bereikt door eerlijk te werken, over hoe doorzettingsvermogen mooi kan zijn. Het merk inspireert door integriteit.
 De combinatie is krachtig omdat ze prestatie sympathiek maakt. Het merk wint hart en hoofd. De valkuil is voorspelbaarheid: de combinatie kan in clichés van overwinning en hoop vervallen. Of de Onschuldige verzacht de Held tot vrijblijvendheid. De combinatie werkt het best wanneer de prestaties echt voor iets staan.`],
  [`hero`, `sage`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Wijze als secundair archetype is iemand die wint omdat hij het beter weet. De Held geeft de prestatie, de Wijze de strategie en het inzicht. Het is een merk dat intelligent vecht. Denk aan high-performance merken die op data leunen, denksporten of consultancy met een meritocratische inslag.
 De toon is gefocust, scherp en zelfverzekerd. Beelden tonen prestatie én analyse, vaak naast elkaar. Boodschappen gaan over wat je kunt bereiken als je strategie en uitvoering combineert. Het merk laat zien dat winnen niet alleen wilskracht is.
 De combinatie is krachtig omdat ze prestatie geloofwaardig maakt. Het merk wordt gerespecteerd op zowel resultaat als methode. De valkuil is kilheid: de combinatie kan menselijkheid verliezen, gericht op resultaat zonder warmte. Of de Wijze remt de Held met overanalyse. De combinatie werkt het best wanneer denken en doen elkaar ontmoeten.`],
  [`hero`, `explorer`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Ontdekker als secundair archetype is iemand die grenzen verlegt door te presteren in onbekend terrein. De Held geeft de prestatie, de Ontdekker geeft het terrein. Het is een merk voor wie de top zoekt buiten de gebaande paden. Denk aan extreme sportmerken, expeditie-uitrusting of pioniers in een vakgebied.
 De toon is gedreven en open. Beelden tonen mensen op uitdagende plekken, in actie, met de horizon als getuige. Boodschappen gaan over wat je ontdekt door je grenzen te verleggen, over wat je leert door te durven. Het merk inspireert door wat het laat zien.
 De combinatie is krachtig omdat ze ambitie en avontuur verenigt. Het merk biedt een grootse versie van zelfontwikkeling. De valkuil is exclusiviteit: de combinatie kan onbereikbaar lijken voor wie geen extreme atleet is. Of de Ontdekker verstrooit de Held en het focus verdwijnt. De combinatie werkt het best wanneer prestatie en verkenning hetzelfde gebaar zijn.`],
  [`hero`, `outlaw`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Rebel als secundair archetype is iemand die wint door tegen de stroom in te gaan. De Held geeft de overwinning, de Rebel geeft de breuk waaruit die voortkomt. Het is een merk voor de underdog die de elite verslaat. Denk aan sportmerken met een outsider-houding of disruptieve uitdagers van marktleiders.
 De toon is strijdvaardig, energiek en eigenzinnig. Beelden tonen mensen die het tegen de verwachting in halen, vaak in ruige settings. Boodschappen gaan over wat je kunt bereiken als je weigert mee te lopen. Het merk inspireert door zijn voorbeeld én door zijn houding.
 De combinatie is krachtig omdat ze ambitie politiek lading geeft. Het merk biedt mensen meer dan resultaat, het biedt een statement. De valkuil is testosterongedreven retoriek: de combinatie kan vervallen in macho-taal of eeuwig oppositioneel blijven. De combinatie werkt het best wanneer winnen en breken in dezelfde beweging zitten.`],
  [`hero`, `magician`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Magiër als secundair archetype is iemand die mensen transformeert door uitdaging. De Held geeft de prestatie en het pushen, de Magiër geeft de verandering die daaruit volgt. Het is een merk dat zegt: word wie je kunt zijn. Denk aan transformatieve fitnessconcepten, krachtige opleidingsmerken of performance-coaches met diepgang.
 De toon is intens, doelgericht en geladen. Beelden tonen mensen vóór en na, of in het moment van doorbraak. Boodschappen gaan over wie je wordt door je grenzen op te zoeken, over de versie aan de andere kant van inspanning. Het merk wijst de weg én daagt uit.
 De combinatie is krachtig omdat ze prestatie betekenis geeft. Het merk biedt meer dan resultaat, het biedt persoonlijke verandering. De valkuil is overheilige retoriek: de combinatie kan vervallen in motivational-cult. Of de Magiër maakt de Held vaag. De combinatie werkt het best wanneer de transformatie concreet en de prestatie echt is.`],
  [`hero`, `lover`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Verleider als secundair archetype is iemand die wint met aantrekkingskracht. De Held geeft de prestatie, de Verleider de stijl en het magnetisme. Het is een merk dat presteren aantrekkelijk maakt. Denk aan sportmerken met een fashion-kant, motormerken met esthetiek of luxe performance-producten.
 De toon is zelfverzekerd, esthetisch en doelgericht. Beelden zijn cinematisch, vaak met aandacht voor lichaam en vorm. Boodschappen gaan over de schoonheid van prestatie, over wat winnen aantrekkelijk maakt. Het merk maakt resultaat begeerlijk.
 De combinatie is krachtig omdat ze prestatie en aantrekkingskracht verbindt. Het merk biedt meer dan succes, het biedt allure. De valkuil is oppervlakkigheid: de combinatie kan vervallen in stijl zonder substantie, of de Verleider kan de Held distraheren van zijn doel. De combinatie werkt het best wanneer de schoonheid uit de prestatie zelf komt.`],
  [`hero`, `jester`, `productieve spanning`, `Een merk met de Held als kern en de Joker als secundair archetype is iemand die wint zonder zichzelf te serieus te nemen. De Held geeft de prestatie, de Joker de relativering. Het is een merk dat ambitie sympathiek maakt. Denk aan teamsporten met humor, of prestatiemerken die hun werk en zichzelf niet plechtig brengen.
 De toon is energiek, opgewekt en zelfbewust. Beelden tonen mensen die hard werken én plezier hebben, vaak met onverwachte details. Boodschappen gaan over winnen op een manier die ook leuk is, over inspanning zonder gewicht. Het merk lacht onderweg.
 De combinatie is krachtig omdat ze ambitie toegankelijk maakt. Het merk biedt prestatie zonder ellende. De valkuil is dat de Joker de Held ondermijnt en winnen onbelangrijk lijkt. Of de Held verdringt de humor en de combinatie wordt forced fun. De combinatie werkt het best wanneer humor en inspanning elkaars natuurlijke aanvulling zijn.`],
  [`hero`, `regular`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Bondgenoot als secundair archetype is iemand die met gewone mensen wint. De Held geeft de prestatie, de Bondgenoot de toegankelijkheid en de gemeenschap. Het is een merk dat zegt: jij kunt dit ook. Denk aan teamsportmerken, hardloopplatformen voor iedereen of fitnessmerken die niet over elite-atleten gaan.
 De toon is bemoedigend, herkenbaar en zonder pretentie. Beelden tonen gewone mensen die iets bereiken, in groep of alleen. Boodschappen gaan over wat haalbaar is voor wie het probeert, over inzet als gemeenschappelijke ervaring. Het merk staat naast zijn gebruikers, niet boven hen.
 De combinatie is krachtig omdat ze prestatie democratiseert. Het merk geeft een breed publiek toegang tot het goede gevoel van presteren. De valkuil is gemiddeldheid: de combinatie kan zo herkenbaar worden dat de inspirerende kracht van de Held verdwijnt. Of de Held maakt de Bondgenoot tot uitsluiting. De combinatie werkt het best wanneer prestatie en gemeenschap elkaar versterken.`],
  [`hero`, `caregiver`, `productieve spanning`, `Een merk met de Held als kern en de Beschermer als secundair archetype is iemand die strijdt om anderen te beschermen. De Held geeft de prestatie en daadkracht, de Beschermer de zorg. Het is een merk dat zegt: ik vecht voor jou. Denk aan defensiemerken, hulporganisaties of merken in veiligheid en zorg die kracht uitstralen.
 De toon is daadkrachtig, betrokken en verantwoordelijk. Beelden tonen mensen die handelen om iets of iemand veilig te stellen. Boodschappen gaan over wat we beschermen door te handelen, over hoe kracht en zorg samenhangen. Het merk laat zien wat het doet en waarom.
 De combinatie is krachtig omdat ze daadkracht een morele lading geeft. Het merk wordt vertrouwd op zijn intentie én resultaat. De valkuil is dramatiek: de combinatie kan vervallen in pathos of helden-retoriek. Of de Beschermer maakt de Held te zacht. De combinatie werkt het best wanneer de zorg concreet en de daad gericht is.`],
  [`hero`, `ruler`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Leider als secundair archetype is iemand die wint en daarmee de standaard wordt. De Held geeft de prestatie, de Leider de positie en het gezag. Het is een merk dat aan de top staat omdat het er hard voor heeft gewerkt. Denk aan toonaangevende sportmerken, marktleiders met een prestatie-erfenis of premium prestatiemerken.
 De toon is zelfverzekerd, gefocust en gezaghebbend. Beelden tonen succes, geschiedenis en uitmuntendheid. Boodschappen gaan over wat het merk heeft bereikt, over de standaard die het zet, over wat het van anderen onderscheidt. Het merk hoeft niet te overtuigen, het laat zijn cijfers spreken.
 De combinatie is krachtig omdat ze succes onderbouwt. Het merk wordt gevolgd omdat het heeft bewezen. De valkuil is arrogantie: de combinatie kan zelfvoldaan worden, of het merk gaat lui worden op zijn positie. De combinatie werkt het best wanneer het leiderschap voortdurend opnieuw wordt verdiend door prestatie.`],
  [`hero`, `creator`, `natuurlijke aanvulling`, `Een merk met de Held als kern en de Creator als secundair archetype is iemand die wint door iets unieks te bouwen. De Held geeft de drive en het doorzettingsvermogen, de Creator de originaliteit en het ambacht. Het is een merk dat presteert door eigen werk. Denk aan ambachtelijke prestatiemerken, sport-design of innovatieve makers in een competitieve markt.
 De toon is gedreven en met aandacht. Beelden tonen werk in uitvoering, processen en resultaten. Boodschappen gaan over wat je bereikt door zelf te bouwen, over het verschil tussen kopiëren en creëren. Het merk presenteert zijn werk als bewijs van zijn ambitie.
 De combinatie is krachtig omdat ze ambitie vormgeeft. Het merk levert wat het belooft, in iets unieks. De valkuil is dat de Held de Creator opjaagt tot snelle resultaten ten koste van het ambacht. Of dat de Creator de Held vertraagt door perfectionisme. De combinatie werkt het best wanneer prestatie en originaliteit elkaars motor zijn.`],
  [`lover`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Verleider als kern en de Onschuldige als secundair archetype is iemand die schoonheid eerlijk maakt. De Verleider geeft de zinnelijkheid, de Onschuldige geeft de zuiverheid. Het is een merk dat verleidt zonder pretenties. Denk aan natuurlijke verzorgingsmerken, eerlijke wijnmerken of zachte hospitality-concepten.
 De toon is warm, esthetisch en oprecht. Beelden zijn mooi maar herkenbaar, intiem maar niet gestileerd. Boodschappen gaan over genot in eenvoud, over schoonheid die je niet hoeft te verdienen. Het merk verleidt zachtjes, zonder druk.
 De combinatie is krachtig omdat ze schoonheid toegankelijk maakt. Het merk biedt genot zonder schuldgevoel. De valkuil is dat de Onschuldige de Verleider afzwakt tot iets te netjes, of dat de Verleider de Onschuldige verdacht maakt. De combinatie werkt het best wanneer de schoonheid uit de eerlijke kern voortkomt.`],
  [`lover`, `sage`, `niet eenvoudig`, `Een merk met de Verleider als kern en de Wijze als secundair archetype is iemand die schoonheid intelligent maakt. De Verleider geeft de aantrekkingskracht, de Wijze de inhoud en de duiding. Het is een merk waarin smaak en kennis samengaan. Denk aan culturele merken, premium magazines of luxe horlogemerken die hun kennis tonen.
 De toon is verzorgd, intelligent en aantrekkelijk. Beelden zijn esthetisch én betekenisvol. Boodschappen gaan over de geschiedenis van een vorm, over de gedachte achter een ontwerp, over schoonheid die uit kennis voortkomt. Het merk biedt verfijning met onderbouwing.
 De combinatie is krachtig omdat ze smaak gefundeerd maakt. Het merk wordt serieus genomen op zowel uiterlijk als inhoud. De valkuil is gemaaktheid: de Wijze kan de Verleider stijf maken, de Verleider kan de Wijze verdacht doen lijken. De combinatie werkt het best wanneer verfijning en inhoud uit hetzelfde fundament komen.`],
  [`lover`, `explorer`, `productieve spanning`, `Een merk met de Verleider als kern en de Ontdekker als secundair archetype is iemand die schoonheid zoekt voorbij de gebaande paden. De Verleider geeft de zinnelijkheid, de Ontdekker geeft de openheid voor het ongepolijste. Het is een merk dat luxe combineert met eigenzinnigheid. Denk aan boutique reizen, niche-parfums of culinaire avonturen.
 De toon is sensorisch en uitnodigend, met een eigen smaak. Beelden zijn esthetisch én onverwacht, mooi én niet-gestileerd. Boodschappen gaan over de schoonheid van het ongewone, over genot voor wie verder kijkt. Het merk verleidt door zijn keuzes.
 De combinatie is krachtig omdat ze luxe een eigen stem geeft. Het merk staat ergens voor, niet alleen ergens voor in stijl. De valkuil is dat de Ontdekker de Verleider doet zoeken naar iets ongepolijst dat geen geloof meer wekt, of dat de Verleider de Ontdekker tam maakt tot stylized luxe. De combinatie werkt het best wanneer de zinnelijkheid uit de plek of het materiaal zelf komt.`],
  [`lover`, `outlaw`, `productieve spanning`, `Een merk met de Verleider als kern en de Rebel als secundair archetype is iemand die schoonheid maakt die de regels breekt. De Verleider geeft de aantrekkingskracht, de Rebel de scherpte en de breuk. Het is een merk dat aantrekkelijk gevaarlijk is. Denk aan modemerken met een edge, parfums die taboes opzoeken of hospitality met een politieke kant.
 De toon is scherp, zinnelijk en niet veilig. Beelden zijn cinematisch, soms ongemakkelijk, vaak provocerend. Boodschappen gaan over verlangen dat geen vergiffenis vraagt, over schoonheid voorbij de norm. Het merk verleidt door tegen de stroom in te gaan.
 De combinatie is krachtig omdat ze aantrekkingskracht en moed verbindt. Het merk wordt gewenst juist omdat het niet braaf is. De valkuil is leeg uitdagen: de combinatie kan vervallen in provocatie zonder substantie. Of de Verleider maakt de Rebel commercieel tam. De combinatie werkt het best wanneer de schoonheid uit een echte breuk voortkomt.`],
  [`lover`, `magician`, `natuurlijke aanvulling`, `Een merk met de Verleider als kern en de Magiër als secundair archetype is iemand die schoonheid omzet in transformatie. De Verleider geeft de zinnelijkheid, de Magiër de belofte van verandering. Het is een merk dat zegt: dit verandert wie je bent. Denk aan premium parfums, luxe wellness of merken die zinnelijk zijn én betekenis dragen.
 De toon is meeslepend, verleidelijk en geladen. Beelden zijn rijk, met een mystieke kant. Boodschappen gaan over wie je wordt door dit te ervaren, over het wonder van het juiste moment. Het merk bouwt een wereld waar je in wilt verdwijnen.
 De combinatie is krachtig omdat ze begeerte en betekenis stapelt. Het merk biedt meer dan een product, het biedt een belofte van transformatie. De valkuil is leegte: de combinatie kan vervallen in gepolijste fantasie zonder substantie. De combinatie werkt het best wanneer de schoonheid daadwerkelijk iets verandert.`],
  [`lover`, `hero`, `natuurlijke aanvulling`, `Een merk met de Verleider als kern en de Held als secundair archetype is iemand die schoonheid en prestatie combineert. De Verleider geeft de aantrekkingskracht, de Held de daadkracht. Het is een merk dat presteren begeerlijk maakt. Denk aan luxe automerken, high-performance fashion of premium sportmerken.
 De toon is zelfverzekerd, esthetisch en doelgericht. Beelden zijn cinematisch, vaak met aandacht voor vorm en prestatie tegelijk. Boodschappen gaan over de schoonheid van het kunnen, over wat aantrekkelijk is aan inzet. Het merk maakt resultaat begeerlijk.
 De combinatie is krachtig omdat ze aantrekkingskracht onderbouwt met prestatie. Het merk wordt gewenst én gerespecteerd. De valkuil is dat de Held de Verleider hard maakt, of dat de Verleider de Held tot stijl reduceert zonder echte prestatie. De combinatie werkt het best wanneer schoonheid en kracht in dezelfde beweging zitten.`],
  [`lover`, `jester`, `natuurlijke aanvulling`, `Een merk met de Verleider als kern en de Joker als secundair archetype is iemand die schoonheid lichtvoetig maakt. De Verleider geeft de zinnelijkheid, de Joker de relativering en de speelsheid. Het is een merk dat genot zonder gewicht biedt. Denk aan modemerken met humor, lifestyle-merken die zichzelf niet serieus nemen of horeca met een knipoog.
 De toon is geestig, esthetisch en uitnodigend. Beelden zijn mooi maar speels, met onverwachte details. Boodschappen gaan over plezier in stijl, over genot dat ook lol mag zijn. Het merk verleidt met een glimlach.
 De combinatie is krachtig omdat ze schoonheid sympathiek maakt. Het merk biedt genot zonder pretentie. De valkuil is dat de Joker de Verleider trivialiseert, dat schoonheid een grap wordt zonder lading. Of de Verleider maakt de Joker overgepolijst. De combinatie werkt het best wanneer humor en schoonheid in dezelfde toon staan.`],
  [`lover`, `regular`, `productieve spanning`, `Een merk met de Verleider als kern en de Bondgenoot als secundair archetype is iemand die schoonheid voor iedereen maakt. De Verleider geeft de zinnelijkheid, de Bondgenoot de toegankelijkheid en herkenbaarheid. Het is een merk dat luxe democratiseert. Denk aan toegankelijke beauty-merken, charcuterie voor thuis of betaalbare hospitality-concepten met smaak.
 De toon is warm, esthetisch en zonder afstand. Beelden tonen mooi maar herkenbaar, mensen zoals jij in mooie momenten. Boodschappen gaan over wat schoonheid kan zijn voor wie het normaal niet voor zichzelf bedoelt, over genot zonder pretentie. Het merk haalt drempels weg.
 De combinatie is krachtig omdat ze genot voor een breed publiek opent. Het merk geeft mensen het gevoel dat ze het verdienen. De valkuil is bleekheid: de Bondgenoot kan de Verleider tam maken, de Verleider kan de Bondgenoot doen overdrijven. De combinatie werkt het best wanneer schoonheid en herkenbaarheid samen kloppen.`],
  [`lover`, `caregiver`, `natuurlijke aanvulling`, `Een merk met de Verleider als kern en de Beschermer als secundair archetype is iemand die schoonheid biedt vanuit zorg. De Verleider geeft de zinnelijkheid, de Beschermer de aandacht en empathie. Het is een merk dat verwennen aanbiedt vanuit oprechte interesse in welzijn. Denk aan luxe wellness, doorvoelde verzorgingsmerken of warme hospitality.
 De toon is warm, intiem en betrokken. Beelden tonen aandacht, zorg en sensorische plekken. Boodschappen gaan over wat je toelaat als iemand om je geeft, over schoonheid als een vorm van zorg. Het merk verleidt door je te zien.
 De combinatie is krachtig omdat ze genot en zorg verbindt. Het merk biedt verwennerij die als oprecht wordt ervaren. De valkuil is dat de Beschermer de Verleider serieus en zwaar maakt, of dat de Verleider de Beschermer vervalst tot een verkooptechniek. De combinatie werkt het best wanneer de zorg uit de schoonheid spreekt.`],
  [`lover`, `ruler`, `natuurlijke aanvulling`, `Een merk met de Verleider als kern en de Leider als secundair archetype is iemand die de standaard van schoonheid bepaalt. De Verleider geeft de aantrekkingskracht, de Leider de positie en het gezag. Het is een merk dat de smaak in zijn categorie zet. Denk aan toonaangevende luxe-merken, premium auto's of high-end horloges.
 De toon is zelfverzekerd, verzorgd en gezaghebbend. Beelden tonen exclusiviteit, traditie en superieure ambacht. Boodschappen gaan over wat schoonheid is volgens dit merk, over de standaard die anderen volgen. Het merk hoeft niet te overtuigen, het zet de norm.
 De combinatie is krachtig omdat ze begeerte koppelt aan autoriteit. Het merk wordt gewenst omdat het de norm is. De valkuil is afstandelijkheid: de combinatie kan zo elitair worden dat ze haar levendigheid verliest. Of zo zelfverzekerd dat ze niet meer evolueert. De combinatie werkt het best wanneer het leiderschap blijft verdiend door de schoonheid.`],
  [`lover`, `creator`, `natuurlijke aanvulling`, `Een merk met de Verleider als kern en de Creator als secundair archetype is iemand die schoonheid maakt met ambacht. De Verleider geeft de aantrekkingskracht, de Creator het ontwerp en de uitvoering. Het is een merk dat zinnelijk én vakkundig is. Denk aan premium designmerken, verfijnde gastronomie of haute couture.
 De toon is verzorgd, esthetisch en met aandacht. Beelden tonen het werk én het proces, materiaal én voltooiing. Boodschappen gaan over hoe schoonheid ontstaat, over ambacht als basis voor genot. Het merk laat zien hoe het tot stand komt.
 De combinatie is krachtig omdat ze schoonheid onderbouwt met ambacht. Het merk wordt gewenst én gerespecteerd. De valkuil is dat de Creator de Verleider serieus en plechtig maakt, of dat de Verleider de Creator tot illustratie reduceert. De combinatie werkt het best wanneer ambacht en zintuiglijke ervaring elkaars vorm zijn.`],
  [`jester`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Joker als kern en de Onschuldige als secundair archetype is iemand die humor brengt vanuit een goedaardige bedoeling. De Joker geeft de speelsheid, de Onschuldige de zuiverheid. Het is een merk dat lacht zonder iemand te raken. Denk aan kindgerichte merken, animaties met een hart of merken die plezier en eerlijkheid combineren.
 De toon is opgewekt, lichtvoetig en warm. Beelden zijn kleurrijk, vaak met een knipoog, maar nooit cynisch. Boodschappen gaan over het plezier van eenvoudige dingen, over lachen als vorm van verbinding. Het merk lacht mét anderen, niet om iets.
 De combinatie is krachtig omdat ze plezier breed toegankelijk maakt. Het merk geeft mensen een glimlach zonder ongemak. De valkuil is bleekheid: de combinatie kan in zoete onschuld vervallen, met humor zonder lading. De combinatie werkt het best wanneer de speelsheid eerlijk is en de eerlijkheid speels.`],
  [`jester`, `sage`, `productieve spanning`, `Een merk met de Joker als kern en de Wijze als secundair archetype is iemand die humor gebruikt om iets uit te leggen. De Joker geeft de speelsheid, de Wijze de inhoud. Het is een merk dat slim genoeg is om zichzelf niet serieus te nemen. Denk aan verklarende media, slimme uitlegplatformen of educatieve merken met een eigen toon.
 De toon is geestig én gefundeerd. Beelden zijn vaak grafisch met een twist. Boodschappen gaan over kennis die niet zwaar hoeft te zijn, over inzichten verpakt in herkenbare humor. Het merk gebruikt grappen om iets duidelijk te maken, niet om iets te ontwijken.
 De combinatie is krachtig omdat ze kennis onthoudbaar maakt. Een grap die iets onthult, blijft hangen. De valkuil is dat de humor de inhoud overschaduwt of de Wijze de Joker afzwakt tot leuk-bedoelde edutainment. De combinatie werkt het best wanneer de grap iets uitlegt en niets verdoezelt.`],
  [`jester`, `explorer`, `natuurlijke aanvulling`, `Een merk met de Joker als kern en de Ontdekker als secundair archetype is iemand die met humor de wereld in trekt. De Joker geeft de speelsheid, de Ontdekker de openheid. Het is een merk dat avontuur lichtvoetig maakt. Denk aan reisvloggers, festivalmerken of avontuurlijke voedselconcepten zonder pretentie.
 De toon is opgewekt, ongedwongen en nieuwsgierig. Beelden tonen mensen die plezier hebben onderweg, met scheve blikken en onverwachte momenten. Boodschappen gaan over avontuur dat geen heldenreis hoeft te zijn, over plezier als reden om te gaan. Het merk lacht onderweg.
 De combinatie is krachtig omdat ze ontdekken laagdrempelig maakt. Het merk neemt zichzelf niet te serieus en haalt anderen daarin mee. De valkuil is oppervlakkigheid: de combinatie kan in feel-good zonder substantie vervallen. De combinatie werkt het best wanneer plezier uit echte ervaring voortkomt.`],
  [`jester`, `outlaw`, `natuurlijke aanvulling`, `Een merk met de Joker als kern en de Rebel als secundair archetype is iemand die met humor het systeem aanvalt. De Joker geeft de grap, de Rebel de strijdlust. Het is een merk dat lacht waar anderen kwaad worden. Denk aan satirische merken, kritische comediakanalen of opstandige merken die ironie als wapen gebruiken.
 De toon is geestig, scherp en soms gemeen. Beelden zijn opvallend, met een grafisch sterke twist. Boodschappen gaan over de absurditeit van regels, over wat we niet hardop zeggen, over taboes met een lach. Het merk gebruikt humor als spiegel.
 De combinatie is krachtig omdat ze verzet aanstekelijk maakt. Het merk wordt gedeeld omdat het lachen oproept én bij blijft. De valkuil is cynisme: de combinatie kan vervallen in spot zonder agenda. Of de Joker zwakt de Rebel af tot vrijblijvende grollen. De combinatie werkt het best wanneer de grap iets blootlegt.`],
  [`jester`, `magician`, `productieve spanning`, `Een merk met de Joker als kern en de Magiër als secundair archetype is iemand die transformatie speels presenteert. De Joker geeft de luchtigheid, de Magiër de visie. Het is een merk dat verwondering met humor brengt. Denk aan creatieve technologie-platformen, avontuurlijke werkplaatsen of magische ervaringen voor wie van plezier houdt.
 De toon is verbeeldingsrijk én lichtvoetig. Beelden zijn speels en suggestief. Boodschappen gaan over wonder dat geen plechtige zaak hoeft te zijn, over fantasie zonder pretentie. Het merk maakt grote ideeën klein toegankelijk.
 De combinatie is krachtig omdat ze magie sympathiek maakt. Het merk biedt verbeelding zonder zwaarte. De valkuil is dat de Joker de Magiër ondermijnt: als alles een grap is, is geen visie meer geloofwaardig. Of de Magiër maakt de Joker plotseling te ernstig. De combinatie werkt het best wanneer plezier de verbeelding versterkt.`],
  [`jester`, `hero`, `productieve spanning`, `Een merk met de Joker als kern en de Held als secundair archetype is iemand die presteert met plezier. De Joker geeft de speelsheid, de Held de drive en de prestatie. Het is een merk dat ambitie sympathiek maakt. Denk aan teamsporten met humor, energieke campagnes met een knipoog of prestatiemerken die zichzelf niet plechtig brengen.
 De toon is energiek, opgewekt en zelfbewust. Beelden tonen mensen die hard werken én plezier hebben. Boodschappen gaan over winnen op een leuke manier, over inspanning zonder gewicht. Het merk lacht onderweg naar de top.
 De combinatie is krachtig omdat ze ambitie toegankelijk maakt. Het merk biedt prestatie zonder ellende. De valkuil is dat de Joker de Held ondermijnt en winnen onbelangrijk lijkt. Of de Held verdringt de humor en de combinatie wordt forced fun. De combinatie werkt het best wanneer humor en inspanning natuurlijk samen lopen.`],
  [`jester`, `lover`, `natuurlijke aanvulling`, `Een merk met de Joker als kern en de Verleider als secundair archetype is iemand die plezier maakt aantrekkelijk. De Joker geeft de luchtigheid, de Verleider de stijl en zinnelijkheid. Het is een merk dat genot speels brengt. Denk aan modemerken met humor, hospitality met een knipoog of cosmetica die plezier en stijl verbindt.
 De toon is geestig, esthetisch en uitnodigend. Beelden zijn mooi maar speels, met onverwachte details. Boodschappen gaan over plezier in stijl, over genot dat ook lol mag zijn. Het merk verleidt met een glimlach.
 De combinatie is krachtig omdat ze schoonheid sympathiek maakt. Het merk biedt genot zonder pretentie. De valkuil is dat de Joker de Verleider trivialiseert, of dat de Verleider de Joker overpolijst tot iets gemaakts. De combinatie werkt het best wanneer humor en schoonheid in dezelfde toon staan.`],
  [`jester`, `regular`, `natuurlijke aanvulling`, `Een merk met de Joker als kern en de Bondgenoot als secundair archetype is iemand die humor maakt voor en met de gewone mens. De Joker geeft de grap, de Bondgenoot de herkenbaarheid en de toegankelijkheid. Het is een merk dat plezier dichtbij brengt. Coolblue is hier een sterke vertegenwoordiger, en veel huiselijke merken volgen.
 De toon is opgewekt, herkenbaar en zonder afstand. Beelden tonen alledaagse situaties met een grap, mensen zoals jij en ik. Boodschappen gaan over kleine ergernissen en kleine vreugdes, met humor die uit het echte leven komt. Het merk spreekt de taal van zijn doelgroep.
 De combinatie is krachtig omdat ze humor democratiseert. Het merk wordt geliefd door zijn herkenbaarheid. De valkuil is platheid: de combinatie kan in al te herkenbare humor vervallen, of zo middlebrow worden dat ze haar onderscheid verliest. De combinatie werkt het best wanneer de humor specifiek is en de herkenbaarheid scherp.`],
  [`jester`, `caregiver`, `niet eenvoudig`, `Een merk met de Joker als kern en de Beschermer als secundair archetype is iemand die humor inzet om mensen op hun gemak te stellen. De Joker geeft de luchtigheid, de Beschermer de zorg. Het is een merk dat moeilijke dingen toegankelijk maakt. Denk aan zorgmerken met een vriendelijke stem, of platforms die met humor over moeilijke onderwerpen praten.
 De toon is warm, lichtvoetig en attent. Beelden zijn vriendelijk, vaak met aandacht voor menselijkheid in lichte zin. Boodschappen gaan over wat moeilijk is, maar in een vorm die ontwapent. Het merk gebruikt humor om iets te dragen, niet om het weg te wuiven.
 De combinatie is krachtig omdat ze ernst lichtvoetig kan maken. Het merk verlaagt drempels. De valkuil is misplaatste humor: de Joker kan de zorg ondermijnen, de Beschermer kan de Joker te terughoudend maken. De combinatie vraagt fijne afstemming, want misslagen worden snel gevoeld.`],
  [`jester`, `ruler`, `niet eenvoudig`, `Een merk met de Joker als kern en de Leider als secundair archetype is iemand die met humor de norm bepaalt. De Joker geeft de speelsheid, de Leider de positie en het gezag. Het is een merk dat zo zelfbewust is dat het zichzelf niet serieus hoeft te nemen. Denk aan iconische merken die met ironie domineren, zoals oudgediende fastfoodmerken die hun eigen positie spelen.
 De toon is geestig, zelfverzekerd en met geheugen voor zijn positie. Beelden tonen succes met een lach, autoriteit met een knipoog. Boodschappen gaan over wat het merk is en hoe het zichzelf relativeert. Het merk leidt zonder zich serieus te willen presenteren.
 De combinatie is krachtig omdat ze leiderschap menselijk maakt. Het merk wordt geliefd én gerespecteerd. De valkuil is dat de Joker het gezag ondermijnt, of dat de Leider de Joker tot een marketing-truc reduceert. De combinatie vraagt een sterke positie als basis, anders werkt de zelfspot niet.`],
  [`jester`, `creator`, `natuurlijke aanvulling`, `Een merk met de Joker als kern en de Creator als secundair archetype is iemand die plezier maakt door te bouwen. De Joker geeft de speelsheid, de Creator het ambacht en de originaliteit. Het is een merk waarin humor en creativiteit hand in hand gaan. Denk aan creatieve studio's met een eigen toon, ontwerpmerken met een knipoog of speelse innovators.
 De toon is verbeeldingsrijk, gemaakt en uitnodigend. Beelden tonen werk met een twist, processen die plezier oproepen. Boodschappen gaan over wat er ontstaat als je speels denkt, over de relatie tussen lol en ambacht. Het merk laat zien dat maken plezier kan zijn.
 De combinatie is krachtig omdat ze creatie aanstekelijk maakt. Het merk inspireert anderen door zijn voorbeeld. De valkuil is dat de Joker het ambacht oppervlakkig laat lijken, of dat de Creator de Joker verstart in conceptueel werk. De combinatie werkt het best wanneer humor de creativiteit voedt en omgekeerd.`],
  [`regular`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Bondgenoot als kern en de Onschuldige als secundair archetype is iemand die toegankelijkheid combineert met eerlijkheid. De Bondgenoot geeft de herkenbaarheid, de Onschuldige geeft de zuiverheid. Het is een merk dat dichtbij staat zonder iets te verbergen. Denk aan vertrouwde huismerken, eerlijke supermarkten of familieconcepten met een lange traditie.
 De toon is helder, vriendelijk en zonder pretentie. Beelden tonen gewone mensen in alledaagse situaties, met een licht en hoopvol licht. Boodschappen gaan over wat het leven goed maakt: samen eten, kleine routines, eerlijke producten. Het merk spreekt geen reclametaal.
 De combinatie is krachtig omdat ze vertrouwen wekt. Het merk voelt als familie. De valkuil is bleekheid: de combinatie kan zo mild worden dat ze geen profiel heeft, een merk dat overal mag zijn maar nergens echt opvalt. De combinatie werkt het best wanneer de eerlijkheid scherp en de toegankelijkheid specifiek is.`],
  [`regular`, `sage`, `natuurlijke aanvulling`, `Een merk met de Bondgenoot als kern en de Wijze als secundair archetype is iemand die kennis deelt op ooghoogte. De Bondgenoot geeft de toegankelijkheid, de Wijze de inhoud. Het is een merk dat slimme dingen op een gewone manier zegt. Denk aan vakbladen voor de gewone professional, educatieve platforms voor breed publiek of adviesmerken die niet boven hun klant staan.
 De toon is helder, vriendelijk en gefundeerd. Beelden tonen mensen, voorbeelden, situaties uit het echte leven. Boodschappen leggen uit met praktische taal, zonder simpel te worden. Het merk vermijdt jargon waar het kan en behoudt diepgang waar het moet.
 De combinatie is krachtig omdat ze kennis democratiseert. Het merk geeft mensen het gevoel dat ze het kunnen begrijpen. De valkuil is afvlakking: te veel toegankelijkheid kan de inhoud uithollen, of de Wijze kan de Bondgenoot opzwepen tot belerend. De combinatie werkt het best wanneer beide dezelfde toon delen.`],
  [`regular`, `explorer`, `natuurlijke aanvulling`, `Een merk met de Bondgenoot als kern en de Ontdekker als secundair archetype is iemand die ontdekken dichtbij brengt. De Bondgenoot geeft de toegankelijkheid, de Ontdekker de horizon. Het is een merk dat avontuur normaal maakt. Denk aan reisorganisaties zonder pretentie, family-friendly outdoor-merken of toegankelijke bushcraft-concepten.
 De toon is open, vriendelijk en zonder afstand. Beelden tonen herkenbare mensen onderweg, niet alleen avonturiers. Boodschappen gaan over wat ontdekken kan zijn voor wie het nog nooit heeft gedaan, over kleine reizen met betekenis. Het merk is gids zonder leraar te zijn.
 De combinatie is krachtig omdat ze ontdekking democratiseert. Het merk geeft een breed publiek toegang tot wat avontuur lijkt. De valkuil is bleekheid: de Bondgenoot kan de Ontdekker ontwapenen tot iets te veiligs. De combinatie werkt het best wanneer toegankelijkheid en horizon elkaar vrij houden.`],
  [`regular`, `outlaw`, `productieve spanning`, `Een merk met de Bondgenoot als kern en de Rebel als secundair archetype is iemand die opstaat namens de gewone mens. De Bondgenoot geeft de verbondenheid, de Rebel geeft de strijd. Het is een merk dat zegt: dit is niet eerlijk, en wij doen er iets aan. Denk aan vakbondsachtige initiatieven, eerlijke prijs-merken of activistische initiatieven met een herkenbaar gezicht.
 De toon is direct, herkenbaar en strijdbaar. Beelden tonen gewone mensen die hun stem verheffen, alledaagse situaties met een politieke laag. Boodschappen gaan over wat oneerlijk is voor mensen zoals jij en ik, over wat we samen kunnen veranderen. Het merk spreekt namens zijn doelgroep.
 De combinatie is krachtig omdat ze rebellie verankert in solidariteit. Het merk wordt geloofd omdat het de pijn van zijn doelgroep deelt. De valkuil is populisme: de combinatie kan in oppervlakkige slogans vervallen. De combinatie werkt het best wanneer de strijd echt voor mensen wordt gevoerd.`],
  [`regular`, `magician`, `productieve spanning`, `Een merk met de Bondgenoot als kern en de Magiër als secundair archetype is iemand die transformatie binnen handbereik brengt. De Bondgenoot geeft de toegankelijkheid, de Magiër de belofte. Het is een merk dat zegt: ook jij. Denk aan democratische technologie, online leerplatforms die mensen iets nieuws laten ontwikkelen, of merken die professionele tools voor iedereen openen.
 De toon is uitnodigend en herkenbaar. Beelden tonen gewone mensen die iets buitengewoons doen. Boodschappen gaan over wat mogelijk wordt voor wie het tot voor kort niet kon, over magie die niet alleen voor ingewijden is. Het merk slecht drempels.
 De combinatie is krachtig omdat ze ambitie democratiseert. Het merk geeft een breed publiek toegang tot wat exclusief leek. De valkuil is bleekheid: de Bondgenoot kan de magie wegvegen, de Magiër kan de Bondgenoot tot een marketing-techniek reduceren. De combinatie werkt het best wanneer de toegankelijkheid de echte verandering niet ondermijnt.`],
  [`regular`, `hero`, `natuurlijke aanvulling`, `Een merk met de Bondgenoot als kern en de Held als secundair archetype is iemand die met gewone mensen prestaties levert. De Bondgenoot geeft de gemeenschap, de Held de drive. Het is een merk dat zegt: jij kunt dit ook. Denk aan teamsportmerken, hardloopplatformen voor iedereen of fitnessmerken die niet over elite-atleten gaan.
 De toon is bemoedigend, herkenbaar en zonder pretentie. Beelden tonen gewone mensen die iets bereiken. Boodschappen gaan over wat haalbaar is voor wie het probeert, over inzet als gemeenschappelijke ervaring. Het merk staat naast zijn gebruikers.
 De combinatie is krachtig omdat ze prestatie democratiseert. Het merk geeft een breed publiek het goede gevoel van presteren. De valkuil is gemiddeldheid: de combinatie kan zo herkenbaar worden dat de inspirerende kracht verdwijnt. De combinatie werkt het best wanneer de prestatie echt als prestatie wordt gevoeld.`],
  [`regular`, `lover`, `productieve spanning`, `Een merk met de Bondgenoot als kern en de Verleider als secundair archetype is iemand die schoonheid voor iedereen biedt. De Bondgenoot geeft de toegankelijkheid, de Verleider de zinnelijkheid. Het is een merk dat luxe democratiseert. Denk aan toegankelijke beauty-merken, charcuterie voor thuis of betaalbare hospitality-concepten.
 De toon is warm, esthetisch en zonder afstand. Beelden tonen mooi maar herkenbaar, mensen zoals jij in mooie momenten. Boodschappen gaan over wat schoonheid kan zijn voor wie het normaal niet voor zichzelf bedoelt. Het merk haalt drempels weg.
 De combinatie is krachtig omdat ze genot voor een breed publiek opent. Het merk geeft mensen het gevoel dat ze het verdienen. De valkuil is dat de Bondgenoot de Verleider tam maakt, of dat de Verleider de Bondgenoot kunstmatig optilt. De combinatie werkt het best wanneer schoonheid en herkenbaarheid samen kloppen.`],
  [`regular`, `jester`, `natuurlijke aanvulling`, `Een merk met de Bondgenoot als kern en de Joker als secundair archetype is iemand die humor maakt voor en met de gewone mens. De Bondgenoot geeft de herkenbaarheid, de Joker de speelsheid. Het is een merk dat plezier dichtbij brengt. Coolblue is een sterke vertegenwoordiger, en veel huiselijke merken volgen.
 De toon is opgewekt, herkenbaar en zonder afstand. Beelden tonen alledaagse situaties met een grap. Boodschappen gaan over kleine ergernissen en kleine vreugdes, met humor uit het echte leven. Het merk spreekt de taal van zijn doelgroep.
 De combinatie is krachtig omdat ze humor democratiseert. Het merk wordt geliefd door zijn herkenbaarheid. De valkuil is platheid: de combinatie kan in al te brede humor vervallen, of zo middlebrow worden dat ze haar onderscheid verliest. De combinatie werkt het best wanneer de humor specifiek en de herkenbaarheid scherp is.`],
  [`regular`, `caregiver`, `natuurlijke aanvulling`, `Een merk met de Bondgenoot als kern en de Beschermer als secundair archetype is iemand die voor de gewone mens zorgt. De Bondgenoot geeft de herkenbaarheid, de Beschermer de zorg en betrouwbaarheid. Het is een merk dat aan jouw kant staat én voor je opkomt. Denk aan zorgverzekeraars met een vriendelijk gezicht, vakbonden of consumentenorganisaties.
 De toon is warm, betrouwbaar en herkenbaar. Beelden tonen gewone mensen in zorgsituaties, met aandacht en respect. Boodschappen gaan over hoe wij voor elkaar zorgen, over wat we samen kunnen oplossen. Het merk is geen autoriteit, het is een metgezel.
 De combinatie is krachtig omdat ze betrouwbaarheid en nabijheid verbindt. Het merk wordt vertrouwd én geliefd. De valkuil is dat de Beschermer de Bondgenoot betuttelend maakt, of dat de Bondgenoot de Beschermer afzwakt tot vrijblijvende vriendelijkheid. De combinatie werkt het best wanneer de zorg concreet en de gelijkwaardigheid voelbaar is.`],
  [`regular`, `ruler`, `productieve spanning`, `Een merk met de Bondgenoot als kern en de Leider als secundair archetype is iemand die de standaard zet door dichtbij te staan. De Bondgenoot geeft de herkenbaarheid, de Leider de positie. Het is een merk dat marktleider werd door de gewone mens te begrijpen. Denk aan grote retailers met een familiekarakter of nationale merken die in elk huishouden te vinden zijn.
 De toon is herkenbaar, zelfverzekerd en zonder afstand. Beelden tonen succes en alledaagsheid in dezelfde frame. Boodschappen gaan over hoe het merk de mensen kent, over de standaard die uit dat begrip is ontstaan. Het merk leidt door te luisteren.
 De combinatie is krachtig omdat ze leiderschap legitimiteit geeft. Het merk wordt gevolgd omdat het de mensen kent. De valkuil is dat de Leider de Bondgenoot doet wegglijden naar elite, of dat de Bondgenoot de Leider niet doet leiden. De combinatie werkt het best wanneer het leiderschap blijft verdiend door de nabijheid.`],
  [`regular`, `creator`, `natuurlijke aanvulling`, `Een merk met de Bondgenoot als kern en de Creator als secundair archetype is iemand die mensen helpt zelf te maken. De Bondgenoot geeft de herkenbaarheid, de Creator het ambacht en de uitdrukking. Het is een merk dat creativiteit democratiseert. Denk aan doe-het-zelf-merken, hobbyplatforms of toegankelijke design-merken voor het gewone huishouden.
 De toon is uitnodigend, praktisch en met aandacht. Beelden tonen mensen aan het werk thuis, processen en resultaten. Boodschappen gaan over wat je kunt maken met wat je hebt, over de waarde van zelfdoen. Het merk maakt creativiteit haalbaar.
 De combinatie is krachtig omdat ze mensen het gevoel geeft dat ze creatief kunnen zijn. Het merk activeert in plaats van te bewonderen. De valkuil is dat de Bondgenoot de Creator tot middelmaat afvlakt, of dat de Creator de Bondgenoot ontoegankelijk maakt door perfectionisme. De combinatie werkt het best wanneer creativiteit echt voor iedereen voelbaar wordt.`],
  [`caregiver`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Beschermer als kern en de Onschuldige als secundair archetype is iemand die zorgt vanuit een puur hart. De Beschermer geeft de zorg en betrokkenheid, de Onschuldige geeft de hoop en de eerlijkheid. Het is een merk dat warm, betrouwbaar en oprecht behulpzaam is. Veel zorgmerken en non-profits leunen tegen deze hoek aan.
 De toon is rustig, geruststellend en empathisch. Beelden tonen menselijke nabijheid, zorg en eenvoudige momenten van verbinding. Boodschappen gaan over er zijn voor elkaar, over kleine gebaren met grote betekenis. Het merk legt geen druk op en moraliseert niet.
 De combinatie is krachtig omdat ze emotionele warmte aan handelen koppelt. Het merk biedt hoop én concrete zorg. De valkuil is sentimentaliteit: te veel zachtheid kan de boodschap krachteloos maken, of het merk overkomen als naïef. De combinatie werkt het best wanneer de zorg concreet en de hoop verdiend is.`],
  [`caregiver`, `sage`, `natuurlijke aanvulling`, `Een merk met de Beschermer als kern en de Wijze als secundair archetype is iemand die zorgt op basis van kennis. De Beschermer geeft de empathie, de Wijze de expertise. Het is een merk dat begrijpt wat er aan de hand is en handelt vanuit verantwoordelijkheid. Denk aan medische merken, juridische adviesbureaus of zorgverzekeraars met een sterke inhoudelijke poot.
 De toon is rustig, deskundig en empathisch. Beelden tonen aandacht en deskundigheid. Boodschappen gaan over begrijpen, voorbereiden en bijstaan. Het merk legt uit waar het ingewikkeld is en stelt gerust waar het kan. Inhoud en zorg zijn één.
 De combinatie is krachtig omdat ze deskundigheid menselijk maakt. Het merk wordt vertrouwd op zowel intentie als inhoud. De valkuil is paternalisme: de Wijze kan belerend worden, de Beschermer betuttelend. De combinatie werkt het best wanneer kennis ondersteunt zonder op te leggen.`],
  [`caregiver`, `explorer`, `productieve spanning`, `Een merk met de Beschermer als kern en de Ontdekker als secundair archetype is iemand die anderen begeleidt op het pad. De Beschermer geeft de zorg, de Ontdekker de openheid. Het is een merk dat avontuur veilig maakt zonder het tam te maken. Denk aan natuurorganisaties, expedities met aandacht of begeleide reizen voor wie nieuw is.
 De toon is bewust, betrokken en wijs. Beelden tonen plekken én mensen die ze betreden, met aandacht voor beide. Boodschappen gaan over hoe je iets ontdekt zonder het kapot te maken, over reizen met respect, over jouw rol in de wereld.
 De combinatie is krachtig omdat ze beleving en moraal koppelt. Het merk laat zien dat avontuur niet ten koste gaat van de wereld. De valkuil is moralisme: de Beschermer kan de Ontdekker schuldgevoel bezorgen. Of de Ontdekker negeert de zorg. De combinatie werkt het best wanneer zorg natuurlijk in de ervaring is verweven.`],
  [`caregiver`, `outlaw`, `productieve spanning`, `Een merk met de Beschermer als kern en de Rebel als secundair archetype is iemand die opkomt voor wie geen stem heeft. De Beschermer geeft de zorg, de Rebel de strijdlust. Het is een merk dat woedend wordt namens een ander. Denk aan dierenwelzijnsorganisaties, kindbeschermingsprogramma's of opstandige sociale ondernemers.
 De toon is fel én warm, scherp én betrokken. Beelden tonen kwetsbaarheid en strijd in dezelfde frame. Boodschappen gaan over wie geen stem heeft, over wat we niet meer accepteren. Het merk vecht uit zorg.
 De combinatie is krachtig omdat ze morele urgentie geeft aan rebellie. De valkuil is sentimentaliteit gekoppeld aan agressie: de combinatie kan zwaar of veroordelend worden. Of de Rebel maakt de Beschermer hard. De combinatie werkt het best wanneer de bescherming concreet en de strijd gericht is.`],
  [`caregiver`, `magician`, `natuurlijke aanvulling`, `Een merk met de Beschermer als kern en de Magiër als secundair archetype is iemand die zorgt voor mensen op een keerpunt. De Beschermer geeft de veiligheid, de Magiër de transformatie. Het is een merk dat groei mogelijk maakt door zorg. Denk aan therapeuten, mind-merken met aandacht of spirituele wellness met diepgang.
 De toon is rustig, warm en geheimzinnig in zachte zin. Beelden tonen mensen die in een veilige ruimte iets ondergaan. Boodschappen gaan over wat je toelaat als je je veilig voelt, over de combinatie van overgave en zorg. Het merk biedt ruimte.
 De combinatie is krachtig omdat ze verandering veilig maakt. Het merk wordt vertrouwd voor wat anders eng zou zijn. De valkuil is zweverigheid: de Magiër kan de Beschermer doen overdrijven in beloftes, of de Beschermer kan de magie verzwakken. De combinatie werkt het best wanneer zorg en visie elkaar versterken.`],
  [`caregiver`, `hero`, `productieve spanning`, `Een merk met de Beschermer als kern en de Held als secundair archetype is iemand die strijdt om anderen te beschermen. De Beschermer geeft de zorg, de Held geeft de prestatie en daadkracht. Het is een merk dat zegt: ik vecht voor jou. Denk aan defensiemerken, hulporganisaties of merken in veiligheid die kracht uitstralen.
 De toon is daadkrachtig, betrokken en verantwoordelijk. Beelden tonen mensen die handelen om iets of iemand veilig te stellen. Boodschappen gaan over wat we beschermen door te handelen, over hoe kracht en zorg samenhangen. Het merk laat zien wat het doet en waarom.
 De combinatie is krachtig omdat ze daadkracht een morele lading geeft. Het merk wordt vertrouwd op zowel intentie als resultaat. De valkuil is dramatiek: de combinatie kan vervallen in pathos of helden-retoriek. De combinatie werkt het best wanneer de zorg concreet en de daad gericht is.`],
  [`caregiver`, `lover`, `natuurlijke aanvulling`, `Een merk met de Beschermer als kern en de Verleider als secundair archetype is iemand die zorgt voor anderen op een sensorische manier. De Beschermer geeft de aandacht, de Verleider geeft de schoonheid en zinnelijkheid. Het is een merk dat verzorgen tot een ervaring maakt. Denk aan luxe wellness, doorvoelde verzorgingsmerken of warme hospitality.
 De toon is warm, intiem en betrokken. Beelden tonen aandacht, zorg en sensorische plekken. Boodschappen gaan over wat je toelaat als iemand om je geeft, over schoonheid als een vorm van zorg. Het merk verleidt door je te zien.
 De combinatie is krachtig omdat ze genot en zorg verbindt. Het merk biedt verwennerij die als oprecht wordt ervaren. De valkuil is dat de Verleider de Beschermer commercialiseert, of dat de Beschermer de Verleider serieus en zwaar maakt. De combinatie werkt het best wanneer de zorg uit de schoonheid spreekt.`],
  [`caregiver`, `jester`, `niet eenvoudig`, `Een merk met de Beschermer als kern en de Joker als secundair archetype is iemand die zorgen oppakt met een glimlach. De Beschermer geeft de aandacht, de Joker de luchtigheid. Het is een merk dat moeilijke dingen toegankelijk maakt zonder ze te bagatelliseren. Denk aan zorgmerken met een vriendelijke stem of platforms die met humor over moeilijke onderwerpen praten.
 De toon is warm, lichtvoetig en attent. Beelden zijn vriendelijk, vaak met aandacht voor menselijkheid in lichte zin. Boodschappen gaan over wat moeilijk is, in een vorm die ontwapent. Het merk gebruikt humor om iets te dragen.
 De combinatie is krachtig omdat ze ernst lichtvoetig kan maken. Het merk verlaagt drempels. De valkuil is misplaatste humor: de Joker kan de zorg ondermijnen, de Beschermer kan de Joker te terughoudend maken. De combinatie vraagt fijne afstemming, want misslagen worden snel gevoeld.`],
  [`caregiver`, `regular`, `natuurlijke aanvulling`, `Een merk met de Beschermer als kern en de Bondgenoot als secundair archetype is iemand die zorgt voor de gewone mens. De Beschermer geeft de zorg, de Bondgenoot de herkenbaarheid. Het is een merk dat aan jouw kant staat én voor je zorgt. Denk aan zorgverzekeraars met een vriendelijk gezicht, vakbonden of consumentenorganisaties.
 De toon is warm, betrouwbaar en herkenbaar. Beelden tonen gewone mensen in zorgsituaties, met aandacht en respect. Boodschappen gaan over hoe wij voor elkaar zorgen, over wat we samen kunnen oplossen. Het merk is geen autoriteit, het is een metgezel.
 De combinatie is krachtig omdat ze betrouwbaarheid en nabijheid verbindt. Het merk wordt vertrouwd én geliefd. De valkuil is dat de Beschermer de Bondgenoot betuttelend maakt, of dat de Bondgenoot de Beschermer afzwakt tot vrijblijvende vriendelijkheid. De combinatie werkt het best wanneer zorg concreet en gelijkwaardigheid voelbaar is.`],
  [`caregiver`, `ruler`, `natuurlijke aanvulling`, `Een merk met de Beschermer als kern en de Leider als secundair archetype is iemand die zorgt vanuit een leidende positie. De Beschermer geeft de toewijding, de Leider de standaard. Het is een merk dat de norm zet voor zorg. Denk aan toonaangevende ziekenhuizen, premium verzekeraars of zorgmerken die zich op kwaliteit profileren.
 De toon is gezaghebbend, warm en zelfverzekerd. Beelden tonen kwaliteit én aandacht, expertise én menselijkheid. Boodschappen gaan over wat zorg op het hoogste niveau betekent, over de standaard die het merk hanteert. Het merk leidt en zorgt tegelijk.
 De combinatie is krachtig omdat ze zorg gezaghebbend maakt. Het merk wordt vertrouwd op zowel positie als intentie. De valkuil is afstandelijkheid: de Leider kan de Beschermer formeel maken, of de Beschermer kan de Leider tot een marketing-kop reduceren. De combinatie werkt het best wanneer kwaliteit en warmte uit hetzelfde fundament komen.`],
  [`caregiver`, `creator`, `natuurlijke aanvulling`, `Een merk met de Beschermer als kern en de Creator als secundair archetype is iemand die zorg vorm geeft. De Beschermer geeft de aandacht, de Creator het ontwerp en de uitvoering. Het is een merk dat zorgt door wat het maakt. Denk aan medische technologie met aandacht voor de gebruiker, of zorgomgevingen die met zorg zijn ontworpen.
 De toon is warm, doordacht en met aandacht voor detail. Beelden tonen menselijke ervaring en de vorm waarin die wordt opgevangen. Boodschappen gaan over hoe een ontwerp het verschil maakt voor wie het gebruikt, over zorg die zichtbaar is in de vorm. Het merk laat zien dat zorg en ambacht samengaan.
 De combinatie is krachtig omdat ze zorg tastbaar maakt. Het merk levert zorg in de vorm van wat het maakt. De valkuil is overontwerp: de Creator kan de Beschermer te clean maken, of de Beschermer kan de Creator beperken tot functioneel ontwerp. De combinatie werkt het best wanneer zorg en ambacht in elkaar overlopen.`],
  [`ruler`, `innocent`, `niet eenvoudig`, `Een merk met de Leider als kern en de Onschuldige als secundair archetype is iemand die autoriteit combineert met eerlijkheid. De Leider geeft de positie en het gezag, de Onschuldige geeft het vertrouwen. Het is een merk dat de standaard zet zonder zijn integriteit te verliezen. Denk aan oudgediende kwaliteitsmerken met een familiekarakter, of nationale instituten met een lange geschiedenis.
 De toon is helder, zelfverzekerd zonder arrogantie, en ingetogen krachtig. Beelden tonen kwaliteit, traditie en bewuste eenvoud. Boodschappen gaan over hoe het hoort, over wat blijft, over standaarden die ergens vandaan komen. Het merk is gezaghebbend zonder afstandelijk te worden.
 De combinatie is krachtig omdat ze geloofwaardigheid en aspiratie verbindt. Het merk wordt gezien als een betrouwbare standaard. De valkuil is dat de Leider de Onschuldige overstemt en de communicatie afstandelijk wordt. Of dat de Onschuldige de Leider verzwakt tot vrijblijvendheid. De combinatie werkt het best wanneer de geschiedenis de claim onderbouwt.`],
  [`ruler`, `sage`, `natuurlijke aanvulling`, `Een merk met de Leider als kern en de Wijze als secundair archetype is iemand die de standaard zet op basis van kennis. De Leider geeft de positie, de Wijze de inhoud. Het is een merk dat de markt vormt door te weten en te bepalen. Denk aan toonaangevende adviesbureaus, normerende instituten of high-end vakmerken.
 De toon is gezaghebbend, doordacht en zelfbewust. Beelden zijn strak, vaak in gedempte of klassieke esthetiek. Boodschappen gaan over wat de norm is, wat de standaard zou moeten zijn. Het merk spreekt zonder te overtuigen, omdat het zijn positie heeft verdiend.
 De combinatie is krachtig omdat ze gezag onderbouwt. Het merk leidt op basis van inhoud, niet alleen op basis van positie. De valkuil is afstandelijkheid en zelfgenoegzaamheid: de combinatie kan zo serieus worden dat ze het contact met de werkelijkheid verliest. De combinatie werkt het best wanneer het leiderschap voortdurend opnieuw wordt verdiend.`],
  [`ruler`, `explorer`, `niet eenvoudig`, `Een merk met de Leider als kern en de Ontdekker als secundair archetype is iemand die marktleider is geworden door pionierswerk. De Leider geeft de positie, de Ontdekker de oorsprong. Het is een merk dat ooit ergens kwam en nu de norm bepaalt. Denk aan premium reismerken met erfgoed of vakgebieden waar ervaring autoriteit wordt.
 De toon is zelfverzekerd, ervaren en kalm. Beelden tonen exclusieve plekken en doorleefde mensen. Boodschappen gaan over wat échte ervaring is, over kwaliteit die je alleen krijgt door tijd. Het merk hoeft niet te overtuigen, het laat zijn track record spreken.
 De combinatie is krachtig omdat ze pioniersgeest tot premium verheft. De valkuil is dat de Leider de Ontdekker stilzet en het merk vastloopt in zijn verleden. Of dat de Ontdekker het gezag ondermijnt door teveel beweging. De combinatie werkt het best wanneer de leider blijft ontdekken en zijn positie blijft verdienen.`],
  [`ruler`, `outlaw`, `productieve spanning`, `Een merk met de Leider als kern en de Rebel als secundair archetype is iemand die zijn macht inzet om dingen anders te doen. De Leider geeft de positie, de Rebel de breuk. Het is een merk dat aan de top staat en die positie gebruikt om te disrupteren. Denk aan grote bedrijven die een sector durven te herzien, of marktleiders die zichzelf opnieuw uitvinden.
 De toon is zelfverzekerd, scherp en met geheugen voor zijn afkomst. Beelden tonen succes met een edge, autoriteit met een stempel. Boodschappen gaan over wat het merk blijft veranderen, over de verantwoordelijkheid van succes. Het merk leidt zonder zich te conformeren.
 De combinatie is krachtig omdat ze macht productief inzet. De valkuil is incoherentie: het merk kan zo gevestigd raken dat de Rebel ongeloofwaardig wordt, of zo blijven veranderen dat zijn leiderschap verdwijnt. De combinatie werkt het best wanneer het leiderschap gebruikt wordt om door te breken, niet om te conserveren.`],
  [`ruler`, `magician`, `natuurlijke aanvulling`, `Een merk met de Leider als kern en de Magiër als secundair archetype is iemand die de toekomst bepaalt vanuit zijn positie. De Leider geeft het gezag, de Magiër de visie. Het is een merk dat zegt: dit is waar de wereld heen gaat, en wij gaan voorop. Denk aan toonaangevende merken in opkomende sectoren of concerns die transformatie bepalen.
 De toon is zelfverzekerd, visionair en gezaghebbend. Beelden tonen toekomst en exclusiviteit. Boodschappen gaan over wat morgen zal zijn, over standaard die nu wordt gezet, over leiderschap als bewijs van visie. Het merk laat geen ruimte voor twijfel.
 De combinatie is krachtig omdat ze ambitie vorm geeft. Het merk dwingt mensen serieus te nemen wat het zegt. De valkuil is hubris: de combinatie kan vervallen in zelfverklaarde profetie zonder grond. De combinatie werkt het best wanneer visie en leiderschap waargemaakt worden door consistente realisatie.`],
  [`ruler`, `hero`, `natuurlijke aanvulling`, `Een merk met de Leider als kern en de Held als secundair archetype is iemand die de standaard zet door te presteren. De Leider geeft de positie, de Held de drive en het resultaat. Het is een merk dat aan de top staat omdat het er hard voor heeft gewerkt. Denk aan toonaangevende sportmerken, marktleiders met een prestatie-erfenis of premium prestatiemerken.
 De toon is zelfverzekerd, gefocust en gezaghebbend. Beelden tonen succes, geschiedenis en uitmuntendheid. Boodschappen gaan over wat het merk heeft bereikt, over de standaard die het zet. Het merk hoeft niet te overtuigen, het laat zijn cijfers spreken.
 De combinatie is krachtig omdat ze succes onderbouwt. Het merk wordt gevolgd omdat het heeft bewezen. De valkuil is arrogantie: de combinatie kan zelfvoldaan worden, of het merk gaat lui worden op zijn positie. De combinatie werkt het best wanneer het leiderschap voortdurend opnieuw wordt verdiend door prestatie.`],
  [`ruler`, `lover`, `natuurlijke aanvulling`, `Een merk met de Leider als kern en de Verleider als secundair archetype is iemand die de standaard van schoonheid bepaalt. De Leider geeft de positie, de Verleider de aantrekkingskracht. Het is een merk dat de smaak in zijn categorie zet. Denk aan toonaangevende luxe-merken, premium auto's of high-end horloges.
 De toon is zelfverzekerd, verzorgd en gezaghebbend. Beelden tonen exclusiviteit, traditie en superieure ambacht. Boodschappen gaan over wat schoonheid is volgens dit merk, over de standaard die anderen volgen. Het merk hoeft niet te overtuigen.
 De combinatie is krachtig omdat ze begeerte aan autoriteit koppelt. Het merk wordt gewenst omdat het de norm is. De valkuil is afstandelijkheid: de combinatie kan zo elitair worden dat ze haar levendigheid verliest. De combinatie werkt het best wanneer het leiderschap blijft verdiend door de schoonheid.`],
  [`ruler`, `jester`, `niet eenvoudig`, `Een merk met de Leider als kern en de Joker als secundair archetype is iemand die met humor de norm bepaalt. De Leider geeft de positie, de Joker de speelsheid. Het is een merk dat zo zelfbewust is dat het zichzelf niet serieus hoeft te nemen. Denk aan iconische merken die met ironie domineren.
 De toon is geestig, zelfverzekerd en met geheugen voor zijn positie. Beelden tonen succes met een lach, autoriteit met een knipoog. Boodschappen gaan over wat het merk is en hoe het zichzelf relativeert. Het merk leidt zonder zich serieus te willen presenteren.
 De combinatie is krachtig omdat ze leiderschap menselijk maakt. Het merk wordt geliefd én gerespecteerd. De valkuil is dat de Joker het gezag ondermijnt, of dat de Leider de Joker tot een marketing-truc reduceert. De combinatie vraagt een sterke positie als basis, anders werkt de zelfspot niet.`],
  [`ruler`, `regular`, `productieve spanning`, `Een merk met de Leider als kern en de Bondgenoot als secundair archetype is iemand die marktleider is geworden door dichtbij de mensen te staan. De Leider geeft de positie, de Bondgenoot de herkenbaarheid. Het is een merk dat groot werd door de gewone mens te begrijpen. Denk aan grote retailers met een familiekarakter of nationale merken in elk huishouden.
 De toon is herkenbaar, zelfverzekerd en zonder afstand. Beelden tonen succes en alledaagsheid in dezelfde frame. Boodschappen gaan over hoe het merk de mensen kent, over de standaard die uit dat begrip is ontstaan. Het merk leidt door te luisteren.
 De combinatie is krachtig omdat ze leiderschap legitimiteit geeft. Het merk wordt gevolgd omdat het de mensen kent. De valkuil is dat de Leider de Bondgenoot doet wegglijden naar elite, of dat de Bondgenoot de Leider niet doet leiden. De combinatie werkt het best wanneer het leiderschap blijft verdiend door de nabijheid.`],
  [`ruler`, `caregiver`, `natuurlijke aanvulling`, `Een merk met de Leider als kern en de Beschermer als secundair archetype is iemand die zorgt vanuit de top. De Leider geeft de positie, de Beschermer de toewijding. Het is een merk dat de norm zet voor zorg en verantwoordelijkheid. Denk aan toonaangevende ziekenhuizen, premium verzekeraars of zorgmerken die zich op kwaliteit profileren.
 De toon is gezaghebbend, warm en zelfverzekerd. Beelden tonen kwaliteit én aandacht, expertise én menselijkheid. Boodschappen gaan over wat zorg op het hoogste niveau betekent, over de standaard die het merk hanteert. Het merk leidt en zorgt tegelijk.
 De combinatie is krachtig omdat ze zorg gezaghebbend maakt. Het merk wordt vertrouwd op zowel positie als intentie. De valkuil is afstandelijkheid: de Leider kan de Beschermer formeel maken, of de Beschermer kan de Leider verzachten tot vrijblijvendheid. De combinatie werkt het best wanneer kwaliteit en warmte uit hetzelfde fundament komen.`],
  [`ruler`, `creator`, `natuurlijke aanvulling`, `Een merk met de Leider als kern en de Creator als secundair archetype is iemand die de standaard zet voor wat gemaakt wordt. De Leider geeft de positie, de Creator het ambacht en de originaliteit. Het is een merk dat marktleider is door zijn werk. Denk aan toonaangevende designhuizen, premium maakmerken of ambachtelijke merken die de norm zijn geworden.
 De toon is zelfverzekerd, verzorgd en met aandacht voor detail. Beelden tonen het werk én de geschiedenis ervan, processen én resultaten. Boodschappen gaan over hoe het merk de standaard zet door wat het maakt. Het merk laat zijn werk spreken.
 De combinatie is krachtig omdat ze leiderschap onderbouwt met ambacht. Het merk wordt gerespecteerd om wat het levert, niet alleen om wie het is. De valkuil is zelfgenoegzaamheid: de combinatie kan vastlopen in een tekst over zijn eigen geschiedenis. Of de Creator wordt geremd door de Leider tot voorzichtige iteratie. De combinatie werkt het best wanneer het maken het leiderschap blijft voeden.`],
  [`creator`, `innocent`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Onschuldige als secundair archetype is iemand die maakt vanuit een eerlijk hart. De Creator geeft het ambacht, de Onschuldige geeft de zuiverheid. Het is een merk dat verfijnd én oprecht is. Denk aan ambachtelijke merken die hun werk met liefde maken, of designers met een ongepretentieuze aanpak.
 De toon is warm en verzorgd. Beelden tonen ambacht, materiaal en eenvoudige schoonheid. Boodschappen gaan over hoe iets gemaakt is, over de mensen erachter, over de kleine keuzes die het verschil maken. Het merk laat zien dat zorgvuldigheid een vorm van eerlijkheid is.
 De combinatie is krachtig omdat ze authenticiteit en kwaliteit verbindt. Het merk straalt zorg en oprechtheid uit. De valkuil is preciositeit: te veel nadruk op vorm kan de eenvoud overschaduwen. De combinatie werkt het best wanneer het ambacht dienstbaar blijft aan een eerlijk product.`],
  [`creator`, `sage`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Wijze als secundair archetype is iemand die maakt op basis van kennis. De Creator geeft het ambacht, de Wijze het inzicht. Het is een merk dat denkt en bouwt. Denk aan researchgedreven designbureaus, technologische pioniers of academische makers.
 De toon is intelligent en gemaakt met aandacht. Beelden zijn vaak verzorgd, met een grafische of conceptuele kant. Boodschappen gaan over hoe een idee vorm krijgt, over de relatie tussen denken en maken. Het merk laat zijn werk zien als bewijs van zijn denken.
 De combinatie is krachtig omdat ze conceptueel werk geloofwaardig maakt. Het merk laat zien dat de vorm volgt uit het denken. De valkuil is dat de Wijze het maken vertraagt door overdenking, of dat de Creator de Wijze tot illustratie reduceert. De combinatie werkt het best wanneer denken en maken hetzelfde proces zijn.`],
  [`creator`, `explorer`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Ontdekker als secundair archetype is iemand die maakt vanuit wat hij tegenkomt. De Creator geeft het ambacht, de Ontdekker de bron. Het is een merk dat zijn werk uit beweging haalt. Denk aan ontwerpmerken die uit reizen geïnspireerd zijn, of foto- en filmstudio's met een sterke handtekening.
 De toon is observerend en gemaakt, met aandacht voor detail én voor het grotere geheel. Beelden tonen plekken, materialen, mensen die werken. Boodschappen gaan over wat een ervaring oplevert, over het ambacht van waarnemen en maken. Het merk laat zijn werk én zijn weg zien.
 De combinatie is krachtig omdat ze beweging en betekenis koppelt. Het merk maakt iets uit wat het tegenkomt. De valkuil is dat de Ontdekker het werk te veel verandert om coherent te blijven, of dat de Creator de Ontdekker afsluit in de werkplaats. De combinatie werkt het best wanneer reis en werk dezelfde ritmiek hebben.`],
  [`creator`, `outlaw`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Rebel als secundair archetype is iemand die maakt om te breken met het bestaande. De Creator geeft het ambacht, de Rebel de tegendraadsheid. Het is een merk dat innovatie als verzet ziet. Denk aan disruptieve designers, alternatieve maakcollectieven of architecten met een politieke agenda.
 De toon is doortastend en doelgericht. Beelden tonen het werk én de breuk waaruit het ontstaat. Boodschappen gaan over waarom het bestaande niet voldoet en wat in plaats daarvan kan komen. Het merk maakt om te tonen wat anders mogelijk is.
 De combinatie is krachtig omdat ze rebellie constructief maakt. Het merk verzet zich met werk, niet alleen met woorden. De valkuil is dat de Rebel het maken vertraagt door alleen te willen breken, of dat de Creator de Rebel afzwakt tot esthetisch experiment. De combinatie werkt het best wanneer breken en bouwen elkaars motor zijn.`],
  [`creator`, `magician`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Magiër als secundair archetype is iemand die maakt om de wereld te veranderen. De Creator geeft het ambacht, de Magiër de visie. Het is een merk dat zijn vorm geeft aan een toekomst. Denk aan visionaire designstudio's of pioniers in technologie die hun eigen wereld creëren.
 De toon is verbeeldingsrijk en doortastend. Beelden tonen werk dat een wereld oproept. Boodschappen gaan over hoe een idee tot leven komt, over de relatie tussen denken en bouwen. Het merk laat zien wat het droomt door wat het maakt.
 De combinatie is krachtig omdat ze visie tastbaar maakt. Het merk levert wat het belooft. De valkuil is overdaad: de combinatie kan vervallen in artistieke pretentie. Of de Magiër verleidt de Creator tot beloftes die het werk niet kan dragen. De combinatie werkt het best wanneer de gemaakte dingen de visie daadwerkelijk dragen.`],
  [`creator`, `hero`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Held als secundair archetype is iemand die wint door iets te bouwen wat anderen niet kunnen. De Creator geeft de originaliteit, de Held de drive en het doorzettingsvermogen. Het is een merk dat presteert door eigen werk. Denk aan ambachtelijke prestatiemerken, sport-design of innovatieve makers in een competitieve markt.
 De toon is gedreven en met aandacht. Beelden tonen werk in uitvoering, processen en resultaten. Boodschappen gaan over wat je bereikt door zelf te bouwen, over het verschil tussen kopiëren en creëren. Het merk presenteert zijn werk als bewijs van zijn ambitie.
 De combinatie is krachtig omdat ze ambitie vormgeeft. Het merk levert wat het belooft, in iets unieks. De valkuil is dat de Held de Creator opjaagt tot snelle resultaten, of dat de Creator de Held vertraagt door perfectionisme. De combinatie werkt het best wanneer prestatie en originaliteit elkaars motor zijn.`],
  [`creator`, `lover`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Verleider als secundair archetype is iemand die schoonheid maakt met ambacht. De Creator geeft het ontwerp, de Verleider de aantrekkingskracht. Het is een merk dat verfijnd én vakkundig is. Denk aan premium designmerken, verfijnde gastronomie of haute couture.
 De toon is verzorgd, esthetisch en met aandacht. Beelden tonen het werk én het proces, materiaal én voltooiing. Boodschappen gaan over hoe schoonheid ontstaat, over ambacht als basis voor genot. Het merk laat zien hoe het tot stand komt.
 De combinatie is krachtig omdat ze schoonheid onderbouwt met ambacht. Het merk wordt gewenst én gerespecteerd. De valkuil is dat de Creator de Verleider serieus en plechtig maakt, of dat de Verleider de Creator tot illustratie reduceert. De combinatie werkt het best wanneer ambacht en zintuiglijke ervaring elkaars vorm zijn.`],
  [`creator`, `jester`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Joker als secundair archetype is iemand die maakt met plezier. De Creator geeft het ambacht, de Joker de speelsheid. Het is een merk waarin creativiteit en humor hand in hand gaan. Denk aan creatieve studio's met een eigen toon, ontwerpmerken met een knipoog of speelse innovators.
 De toon is verbeeldingsrijk, gemaakt en uitnodigend. Beelden tonen werk met een twist, processen die plezier oproepen. Boodschappen gaan over wat er ontstaat als je speels denkt, over de relatie tussen lol en ambacht. Het merk laat zien dat maken plezier kan zijn.
 De combinatie is krachtig omdat ze creatie aanstekelijk maakt. Het merk inspireert anderen door zijn voorbeeld. De valkuil is dat de Joker het ambacht oppervlakkig laat lijken, of dat de Creator de Joker verstart. De combinatie werkt het best wanneer humor de creativiteit voedt en omgekeerd.`],
  [`creator`, `regular`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Bondgenoot als secundair archetype is iemand die maakt voor en met de gewone mens. De Creator geeft het ambacht, de Bondgenoot de toegankelijkheid. Het is een merk dat creativiteit democratiseert. Denk aan doe-het-zelf-merken, hobbyplatforms of toegankelijke design-merken voor het gewone huishouden.
 De toon is uitnodigend, praktisch en met aandacht. Beelden tonen mensen aan het werk thuis, processen en resultaten. Boodschappen gaan over wat je kunt maken met wat je hebt, over de waarde van zelfdoen. Het merk maakt creativiteit haalbaar.
 De combinatie is krachtig omdat ze mensen het gevoel geeft dat ze creatief kunnen zijn. Het merk activeert in plaats van te bewonderen. De valkuil is dat de Bondgenoot de Creator tot middelmaat afvlakt, of dat de Creator de Bondgenoot ontoegankelijk maakt door perfectionisme. De combinatie werkt het best wanneer creativiteit echt voor iedereen voelbaar wordt.`],
  [`creator`, `caregiver`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Beschermer als secundair archetype is iemand die maakt vanuit zorg. De Creator geeft het ambacht, de Beschermer de aandacht voor de gebruiker. Het is een merk dat zorgvuldig ontwerpt voor wie het zal gebruiken. Denk aan medische technologie met aandacht voor de gebruiker, of zorgomgevingen die met zorg zijn ontworpen.
 De toon is warm, doordacht en met aandacht voor detail. Beelden tonen menselijke ervaring en de vorm waarin die wordt opgevangen. Boodschappen gaan over hoe een ontwerp het verschil maakt voor wie het gebruikt. Het merk laat zien dat zorg en ambacht samengaan.
 De combinatie is krachtig omdat ze zorg tastbaar maakt. Het merk levert zorg in de vorm van wat het maakt. De valkuil is overontwerp: de Creator kan de Beschermer te clean maken, of de Beschermer kan de Creator beperken tot functioneel ontwerp. De combinatie werkt het best wanneer zorg en ambacht in elkaar overlopen.`],
  [`creator`, `ruler`, `natuurlijke aanvulling`, `Een merk met de Creator als kern en de Leider als secundair archetype is iemand die de standaard zet door zijn werk. De Creator geeft het ambacht, de Leider de positie. Het is een merk dat marktleider is geworden door wat het maakt. Denk aan toonaangevende designhuizen, premium maakmerken of ambachtelijke merken die de norm zijn geworden.
 De toon is zelfverzekerd, verzorgd en met aandacht voor detail. Beelden tonen het werk én de geschiedenis ervan. Boodschappen gaan over hoe het merk de standaard zet door wat het maakt. Het merk laat zijn werk spreken.
 De combinatie is krachtig omdat ze leiderschap onderbouwt met ambacht. Het merk wordt gerespecteerd om wat het levert. De valkuil is zelfgenoegzaamheid: de combinatie kan vastlopen in een tekst over zijn eigen geschiedenis. Of de Creator wordt geremd door de Leider tot voorzichtige iteratie. De combinatie werkt het best wanneer het maken het leiderschap blijft voeden.`]
];

