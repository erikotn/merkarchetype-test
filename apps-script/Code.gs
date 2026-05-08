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
  // Naar nested object: { primary: { secondary: description } }
  const out = {};
  rows.forEach(r => {
    const p = String(r.primary || '').trim();
    const s = String(r.secondary || '').trim();
    const d = String(r.description || '').trim();
    if (!p || !s || !d) return;
    if (!out[p]) out[p] = {};
    out[p][s] = d;
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

function SEED_DATA() {
  return {
    archetypes: JSON.parse(ARCHETYPES_SEED),
    questions: JSON.parse(QUESTIONS_SEED),
    mappings: JSON.parse(MAPPINGS_SEED),
    combinations: JSON.parse(COMBINATIONS_SEED)
  };
}

const ARCHETYPES_SEED = '[["key", "name", "emoji", "subtitle", "color", "promise", "description", "communication", "traits", "brands"], ["innocent", "De Onschuldige", "🌼", "puur, optimistisch, eenvoudig", "#7AC74F", "Het komt allemaal goed.", "De Onschuldige staat voor puurheid, optimisme en eenvoud. Dit archetype gelooft in het goede en wil zowel zichzelf als anderen gelukkig maken. Het is een betrouwbaar, eerlijk en nostalgisch merk dat veiligheid en positiviteit uitstraalt.", "De Onschuldige communiceert op een eenvoudige, heldere en positieve manier. De toon is optimistisch en geruststellend, zonder cynisme of complexiteit. Beeldgebruik is clean, licht en vaak nostalgisch. Boodschappen zijn rechtdoorzee en bieden troost en zekerheid. Humor is vriendelijk en onschuldig. Dit merk spreekt over waarden als eerlijkheid, betrouwbaarheid en traditionele kwaliteit.", "Positiviteit, vertrouwen, eenvoud, nostalgie", "Campina, Friesche Vlag, Verkade, Dove, Coca-Cola"], ["explorer", "De Ontdekker", "🧭", "vrijheid, avontuur, ontdekking", "#9CCC65", "Ontdek jezelf en de wereld.", "De Ontdekker staat voor vrijheid, avontuur en ontdekking. Dit archetype wil grenzen verleggen, nieuwe ervaringen opdoen en authentiek zijn. Het is onafhankelijk, moedig en pionierend, altijd op zoek naar wat er verder nog te ontdekken valt.", "De Ontdekker communiceert inspirerend en bevrijdend. De toon is avontuurlijk en nodigt uit tot actie. Beelden tonen wijde landschappen, reizen en grenzeloosheid. De boodschap gaat over zelfontdekking, individualiteit en het verleggen van grenzen. Dit merk spreekt mensen aan die zich niet willen laten beperken en hun eigen pad willen bewandelen. Verhalen gaan over expedities, ontdekkingen en persoonlijke groei.", "Avontuur, autonomie, authenticiteit, ontdekking", "KLM, ANWB, Staatsloterij, The North Face, Jeep"], ["sage", "De Wijze", "🦉", "kennis, expertise, betrouwbaarheid", "#388E3C", "De waarheid zal je vrijmaken.", "De Wijze staat voor kennis, waarheid en intelligentie. Dit archetype streeft ernaar de wereld te begrijpen en die kennis te delen met anderen. Het is een expert, een betrouwbare bron van informatie die analytisch en doordacht te werk gaat.", "De Wijze communiceert helder, informatief en onderbouwd. De toon is educatief en gezaghebbend, maar niet betuttelend. Feiten, onderzoek en expertise staan centraal. De communicatie daagt uit om na te denken en nodigt uit tot leren. Beeldmateriaal is vaak strak en professioneel. Jargon is toegestaan als het relevant is. Dit merk positioneert zich als thought leader en betrouwbare kennisbron.", "Wijsheid, expertise, analyse, waarheid", "NOS, NU.nl, Douwe Egberts, BBC, Google"], ["hero", "De Held", "🦸", "moed, prestatie, overwinning", "#E53935", "Waar een wil is, is een weg.", "De Held staat voor moed, prestatie en overwinning. Dit archetype wil impact maken en uitdagingen overwinnen. Het is dapper, sterk en gedreven naar uitmuntendheid.", "De Held communiceert motiverend, krachtig en doelgericht. De toon is inspirerend en roept op tot actie en prestatie. Beelden tonen overwinning, doorzettingsvermogen en succes. De boodschap is dat je alles kunt bereiken als je maar hard genoeg werkt. Dit merk daagt mensen uit het beste uit zichzelf te halen. Verhalen gaan over overwinnen van obstakels, competitie en triomf. De communicatie is energiek en vol zelfvertrouwen.", "Moed, doorzettingsvermogen, prestatie, kracht", "Jumbo, Rabobank, Heineken, Nike, Adidas, FedEx"], ["outlaw", "De Rebel", "🔥", "regelbreker, anders, disruptief", "#F57C00", "De regels zijn er om gebroken te worden.", "De Rebel staat voor rebellie, bevrijding en revolutie. Dit archetype daagt de status quo uit en durft te provoceren. Het is radicaal, disruptief en authentiek rebels.", "De Rebel communiceert provocerend, direct en zonder compromissen. De toon is rebels en uitdagend, soms zelfs confronterend. Beeldgebruik is ruig, edgy en onconventioneel. Dit merk spreekt taboes aan en kiest bewust positie. De communicatie roept op tot actie en verandering. Humor kan cynisch of sarcastisch zijn. Dit merk trekt mensen aan die zich niet conformeren en anders durven te zijn.", "Rebellie, bevrijding, revolutie, provocatie", "G-Star RAW, Tony\'s Chocolonely, Transavia, Harley-Davidson, Diesel, Virgin"], ["magician", "De Magiër", "✨", "transformatie, visie, wonder", "#FB8C00", "Alles is mogelijk.", "De Magiër staat voor transformatie, visie en het waar maken van dromen. Dit archetype inspireert door ogenschijnlijk onmogelijke dingen mogelijk te maken. Het is charismatisch, visionair en transformerend.", "De Magiër communiceert mysterieus, inspirerend en visionair. De toon is magisch en belooft transformatie. Beelden zijn vaak sprookjesachtig, verbeeldingsvol of toekomstgericht. De communicatie gaat over dromen die werkelijkheid worden, over magie en wonder. Dit merk toont hoe het het leven van mensen kan transformeren. Verhalen zijn meeslepend en emotioneel. De boodschap is dat met dit merk alles binnen bereik ligt.", "Transformatie, visie, inspiratie, innovatie", "Efteling, Philips, Disney, Tesla"], ["regular", "De Bondgenoot", "🤝", "erbij horen, toegankelijk, realisme", "#7E57C2", "Wij begrijpen je.", "De Bondgenoot staat voor verbinding, eerlijkheid en toegankelijkheid. Dit archetype wil erbij horen en echte connectie maken met anderen. Het is betrouwbaar, down-to-earth en herkenbaar.", "De Bondgenoot communiceert toegankelijk, herkenbaar en zonder poespas. De toon is vriendelijk, eerlijk en praktisch. Beelden tonen gewone mensen in alledaagse situaties. De communicatie is realistisch en zonder pretentie. Dit merk spreekt de taal van de doelgroep en toont begrip voor hun uitdagingen. Verhalen gaan over herkenbare situaties en praktische oplossingen. De boodschap is: wij zijn net als jij en staan aan jouw kant.", "Toegankelijkheid, eerlijkheid, realisme, verbinding", "Albert Heijn, HEMA, Gamma, Kruidvat, IKEA, Levi\'s, eBay"], ["lover", "De Verleider", "🌹", "intimiteit, plezier, sensualiteit", "#EC407A", "Geniet van het moment.", "De Verleider staat voor passie, intimiteit en schoonheid. Dit archetype streeft naar nabijheid, genot en betekenisvolle ervaringen. Het is sensueel, warm en gepassioneerd.", "De Verleider communiceert sensueel, emotioneel en intiem. De toon is warm, verleidelijk en persoonlijk. Beelden zijn esthetisch, romantisch en appelleren aan de zintuigen. De communicatie gaat over genieten, verwennen en het ervaren van schoonheid. Dit merk spreekt over kwaliteitstijd, zintuiglijke ervaringen en het belang van intimiteit en verbinding. Verhalen zijn emotioneel geladen en persoonlijk. De boodschap is dat je het verdient om jezelf te verwennen.", "Passie, intimiteit, schoonheid, sensualiteit", "Rituals, De Bijenkorf, Spa, Godiva, Chanel, Alfa Romeo"], ["jester", "De Joker", "🎭", "plezier, humor, luchtigheid", "#AB47BC", "Het leven is een feest.", "De Joker staat voor plezier, humor en luchtigheid. Dit archetype wil het leven vieren en anderen aan het lachen maken. Het is speels, optimistisch en spontaan.", "De Joker communiceert humoristisch, energiek en ongedwongen. De toon is lichtvoetig, grappig en entertainend. Beelden zijn kleurrijk, speels en onverwacht. De communicatie durft gek te doen en neemt zichzelf niet te serieus. Dit merk gebruikt humor, woordgrappen en onverwachte wendingen. Verhalen zijn vermakelijk en maken mensen aan het lachen. De boodschap is dat het leven te kort is om je zorgen te maken en dat je vooral moet genieten.", "Humor, plezier, spontaniteit, luchtigheid", "Coolblue, Ben, bol.com, Old Spice, M&M\'s, Skittles"], ["caregiver", "De Beschermer", "🛡️", "zorg, bescherming, medeleven", "#5C6BC0", "Ik zorg voor je.", "De Beschermer staat voor zorg, bescherming en compassie. Dit archetype wil anderen helpen en beschermen. Het is warm, ondersteunend en altruïstisch.", "De Beschermer communiceert empathisch, geruststellend en ondersteunend. De toon is zorgzaam, warm en betrokken. Beelden tonen zorg, bescherming en menselijke verbinding. De communicatie draait om het welzijn van de ander en toont begrip voor zorgen en behoeften. Dit merk stelt zich dienstbaar op en spreekt over verantwoordelijkheid en toewijding. Verhalen gaan over het verschil maken in iemands leven en er zijn wanneer het nodig is. De boodschap is dat je erop kunt vertrouwen dat dit merk voor je zorgt.", "Zorg, bescherming, generositeit, empathie", "Zilveren Kruis, CZ, Rode Kruis, VGZ, Unicef, Volvo, Johnson & Johnson"], ["ruler", "De Leider", "👑", "controle, leiderschap, status", "#1E88E5", "Macht creëert mogelijkheden.", "De Leider staat voor controle, leiderschap en succes. Dit archetype streeft naar het creëren van welvaart en stabiliteit. Het is dominant, georganiseerd en exclusief.", "De Leider communiceert zelfverzekerd, autoritair en prestigieus. De toon is formeel, krachtig en gebiedend. Beelden tonen luxe, succes en exclusiviteit. De communicatie gaat over leiderschap, controle en het bereiken van de top. Dit merk positioneert zich als marktleider en statusmerk. Verhalen gaan over excellentie, erfgoed en het bij de elite horen. De boodschap is dat dit merk voor winners is, voor mensen die het beste verdienen en het beste willen.", "Leiderschap, controle, status, stabiliteit", "ABN AMRO, ING, Nationale Nederlanden, Mercedes-Benz, Rolex, Microsoft"], ["creator", "De Creator", "🎨", "innovatie, creativiteit, verbeelding", "#26C6DA", "Als je het kunt bedenken, kun je het maken.", "De Creator staat voor creativiteit, innovatie en verbeelding. Dit archetype wil iets waardevols en blijvends creëren. Het is creatief, innovatief en authentiek non-conform.", "De Creator communiceert inspirerend, origineel en kunstzinnig. De toon is creatief, innovatief en soms onconventioneel. Beelden zijn vaak artistiek, vernieuwend of design-gedreven. De communicatie gaat over zelfexpressie, creativiteit en het benutten van je verbeeldingskracht. Dit merk moedigt aan om dingen te maken, te ontwerpen en jezelf uit te drukken. Verhalen gaan over het creatieve proces, innovatie en het tot leven brengen van ideeën. De boodschap is dat iedereen creatief kan zijn en dat dit merk de tools biedt om je visie te realiseren.", "Creativiteit, innovatie, verbeelding, expressie", "ASML, Coolblue, Apple, Lego, Adobe"]]';

const QUESTIONS_SEED = '[["id","left","right"],[1,"We zijn speels en spontaan","We zijn serieus en verantwoordelijk"],[2,"We streven naar vrijheid en avontuur","We streven naar veiligheid en zekerheid"],[3,"Onze toon is vriendelijk en zorgzaam","Onze toon is sterk en beschermend"],[4,"We willen de wereld verbeteren","We willen het leven gemakkelijker maken"],[5,"We staan voor authenticiteit en echtheid","We staan voor succes en prestige"],[6,"We inspireren door verbeeldingskracht","We inspireren door logica en efficiëntie"],[7,"We zoeken verandering en revolutie","We zoeken stabiliteit en continuïteit"],[8,"Onze communicatie is warm en menselijk","Onze communicatie is cool en professioneel"],[9,"We richten ons op gemeenschap en verbinding","We richten ons op prestatie en status"],[10,"We helpen mensen hun potentieel te zien","We helpen mensen zich veilig te voelen"],[11,"We geloven in dromen en magie","We geloven in feiten en realiteit"],[12,"We willen dat klanten zich vrij voelen","We willen dat klanten zich beschermd weten"],[13,"We spreken als een vriend","We spreken als een expert"],[14,"We moedigen avontuur en ontdekking aan","We moedigen orde en structuur aan"],[15,"We brengen originaliteit","We brengen betrouwbaarheid"],[16,"We streven naar perfectie","We streven naar authenticiteit"],[17,"We staan voor discipline en beheersing","We staan voor passie en emotie"],[18,"We willen anderen inspireren","We willen anderen dienen"],[19,"We waarderen speelsheid","We waarderen wijsheid"],[20,"We geloven in rebellie","We geloven in traditie"],[21,"Onze klanten zien ons als een gids","Onze klanten zien ons als een bondgenoot"],[22,"We zijn visionair en idealistisch","We zijn praktisch en realistisch"],[23,"We streven naar groei en zelfontplooiing","We streven naar comfort en voorspelbaarheid"],[24,"We spreken met humor","We spreken met ernst"],[25,"We positioneren ons als uitdager","We positioneren ons als leider"],[26,"We draaien om emotie","We draaien om ratio"],[27,"We zijn verleidelijk en aantrekkelijk","We zijn nuchter en functioneel"],[28,"We willen anderen verrassen","We willen anderen geruststellen"],[29,"We bouwen ons merk op intuïtie","We bouwen ons merk op analyse"],[30,"We streven naar vrijheid van expressie","We streven naar controle en orde"]]';

const MAPPINGS_SEED = '[["question_id","answer","archetype","points"],[1,1,"jester",2],[1,1,"explorer",1],[1,2,"jester",1],[1,2,"explorer",1],[1,4,"ruler",1],[1,4,"caregiver",1],[1,5,"ruler",2],[1,5,"caregiver",2],[2,1,"explorer",2],[2,1,"outlaw",1],[2,2,"explorer",1],[2,2,"hero",1],[2,4,"innocent",1],[2,4,"caregiver",1],[2,5,"innocent",2],[2,5,"caregiver",2],[3,1,"caregiver",2],[3,1,"lover",1],[3,2,"caregiver",1],[3,2,"regular",1],[3,4,"hero",1],[3,4,"ruler",1],[3,5,"hero",2],[3,5,"ruler",1],[4,1,"hero",2],[4,1,"magician",1],[4,1,"sage",1],[4,2,"hero",1],[4,2,"magician",1],[4,4,"regular",1],[4,4,"caregiver",1],[4,5,"regular",2],[4,5,"caregiver",2],[5,1,"regular",2],[5,1,"explorer",1],[5,1,"innocent",1],[5,2,"regular",1],[5,2,"explorer",1],[5,4,"ruler",1],[5,4,"hero",1],[5,5,"ruler",2],[5,5,"hero",1],[6,1,"creator",2],[6,1,"magician",2],[6,2,"creator",1],[6,2,"magician",1],[6,4,"sage",1],[6,4,"ruler",1],[6,5,"sage",2],[6,5,"ruler",1],[7,1,"outlaw",2],[7,1,"magician",1],[7,1,"hero",1],[7,2,"outlaw",1],[7,2,"magician",1],[7,4,"innocent",1],[7,4,"ruler",1],[7,4,"caregiver",1],[7,5,"innocent",2],[7,5,"ruler",1],[7,5,"caregiver",1],[8,1,"caregiver",2],[8,1,"lover",1],[8,1,"regular",1],[8,2,"caregiver",1],[8,2,"regular",1],[8,4,"sage",1],[8,4,"ruler",1],[8,5,"sage",2],[8,5,"ruler",1],[9,1,"regular",2],[9,1,"caregiver",2],[9,2,"regular",1],[9,2,"caregiver",1],[9,4,"hero",1],[9,4,"ruler",1],[9,5,"hero",2],[9,5,"ruler",2],[10,1,"magician",2],[10,1,"hero",1],[10,1,"sage",1],[10,2,"magician",1],[10,2,"sage",1],[10,4,"caregiver",1],[10,4,"innocent",1],[10,5,"caregiver",2],[10,5,"innocent",2],[11,1,"magician",2],[11,1,"creator",1],[11,1,"lover",1],[11,2,"magician",1],[11,2,"creator",1],[11,4,"sage",1],[11,4,"regular",1],[11,5,"sage",2],[11,5,"regular",2],[12,1,"explorer",2],[12,1,"outlaw",1],[12,2,"explorer",1],[12,2,"hero",1],[12,4,"caregiver",1],[12,4,"innocent",1],[12,5,"caregiver",2],[12,5,"innocent",2],[13,1,"regular",2],[13,1,"jester",1],[13,2,"regular",1],[13,2,"caregiver",1],[13,4,"sage",1],[13,4,"ruler",1],[13,5,"sage",2],[13,5,"ruler",1],[14,1,"explorer",2],[14,1,"hero",1],[14,2,"explorer",1],[14,2,"creator",1],[14,4,"ruler",1],[14,4,"innocent",1],[14,5,"ruler",2],[14,5,"sage",1],[15,1,"creator",2],[15,1,"outlaw",1],[15,1,"magician",1],[15,2,"creator",1],[15,2,"outlaw",1],[15,4,"regular",1],[15,4,"innocent",1],[15,4,"caregiver",1],[15,5,"regular",2],[15,5,"innocent",2],[15,5,"caregiver",1],[16,1,"ruler",2],[16,1,"hero",1],[16,1,"sage",1],[16,2,"ruler",1],[16,2,"hero",1],[16,4,"regular",1],[16,4,"explorer",1],[16,5,"regular",2],[16,5,"explorer",2],[17,1,"ruler",2],[17,1,"sage",1],[17,2,"ruler",1],[17,2,"hero",1],[17,4,"lover",1],[17,4,"outlaw",1],[17,4,"jester",1],[17,5,"lover",2],[17,5,"outlaw",1],[17,5,"jester",1],[18,1,"hero",2],[18,1,"magician",1],[18,1,"sage",1],[18,2,"hero",1],[18,2,"magician",1],[18,4,"caregiver",1],[18,4,"regular",1],[18,5,"caregiver",2],[18,5,"regular",2],[19,1,"jester",2],[19,1,"innocent",1],[19,2,"jester",1],[19,2,"lover",1],[19,4,"sage",1],[19,4,"ruler",1],[19,5,"sage",2],[19,5,"ruler",1],[20,1,"outlaw",2],[20,1,"creator",1],[20,2,"outlaw",1],[20,2,"hero",1],[20,4,"innocent",1],[20,4,"ruler",1],[20,4,"caregiver",1],[20,5,"innocent",2],[20,5,"ruler",1],[20,5,"caregiver",1],[21,1,"sage",2],[21,1,"hero",1],[21,1,"magician",1],[21,2,"sage",1],[21,2,"hero",1],[21,4,"regular",1],[21,4,"caregiver",1],[21,5,"regular",2],[21,5,"caregiver",2],[22,1,"magician",2],[22,1,"hero",1],[22,1,"creator",1],[22,2,"magician",1],[22,2,"hero",1],[22,4,"regular",1],[22,4,"sage",1],[22,5,"regular",2],[22,5,"sage",1],[23,1,"explorer",2],[23,1,"hero",1],[23,1,"magician",1],[23,2,"explorer",1],[23,2,"hero",1],[23,4,"innocent",1],[23,4,"caregiver",1],[23,5,"innocent",2],[23,5,"caregiver",2],[24,1,"jester",2],[24,1,"lover",1],[24,2,"jester",1],[24,2,"regular",1],[24,4,"sage",1],[24,4,"ruler",1],[24,4,"hero",1],[24,5,"sage",2],[24,5,"ruler",1],[24,5,"hero",1],[25,1,"outlaw",2],[25,1,"hero",1],[25,2,"outlaw",1],[25,2,"hero",1],[25,4,"ruler",1],[25,4,"sage",1],[25,5,"ruler",2],[25,5,"sage",1],[26,1,"lover",2],[26,1,"jester",1],[26,1,"caregiver",1],[26,2,"lover",1],[26,2,"caregiver",1],[26,4,"sage",1],[26,4,"ruler",1],[26,5,"sage",2],[26,5,"ruler",2],[27,1,"lover",2],[27,1,"magician",1],[27,2,"lover",1],[27,2,"creator",1],[27,4,"regular",1],[27,4,"sage",1],[27,5,"regular",2],[27,5,"sage",1],[28,1,"magician",2],[28,1,"jester",1],[28,1,"outlaw",1],[28,2,"magician",1],[28,2,"creator",1],[28,4,"caregiver",1],[28,4,"innocent",1],[28,5,"caregiver",2],[28,5,"innocent",2],[29,1,"creator",2],[29,1,"magician",1],[29,1,"lover",1],[29,2,"creator",1],[29,2,"magician",1],[29,4,"sage",1],[29,4,"ruler",1],[29,5,"sage",2],[29,5,"ruler",2],[30,1,"creator",2],[30,1,"outlaw",1],[30,1,"explorer",1],[30,2,"creator",1],[30,2,"outlaw",1],[30,4,"ruler",1],[30,4,"sage",1],[30,5,"ruler",2],[30,5,"sage",1]]';
const COMBINATIONS_SEED = '[["primary", "secondary", "description"], ["hero", "sage", "Een merk dat moed combineert met expertise. Klanten zien jullie als een dappere gids: iemand die kennis heeft én durft te handelen. Het beste van twee werelden — vol zelfvertrouwen en onderbouwd."]]';

