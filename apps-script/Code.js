/**
 * BUDGET · API du budget de Devlin (Apps Script) · v5.15
 * ------------------------------------------
 * Relié au fichier budget qui le contient : lit Paramètres, Échéancier, Journal et « À venir », et y écrit
 * ce que l'appli envoie. Chaque demande doit contenir la clé (menu Budget › Choisir ma clé).
 * Points clés : envois rejouables sans doublon après une coupure (uid gardé 30 jours, propriété UIDS) ;
 * plusieurs envois en un « lot » sous un seul verrou ; fichier occupé → OCCUPE (l'appli réessaie) ;
 * onglet caché « Diagnostic » (une ligne par demande) ; libellés écrits en texte brut.
 * Colonne « Payé en » du Journal : « Espèces » ou « En ligne ». Préférences de l'appli gardées dans les propriétés du script.
 *
 * Installation : menu Budget › Choisir ma clé, puis Déployer › Nouveau déploiement › Application Web
 * (Exécuter en tant que : moi · Accès : tout le monde). Historique détaillé : git du dépôt 02_Suivi_budget_moi.
 *
 * @OnlyCurrentDoc
 * Le script ne demande l'accès qu'au fichier qui le contient, rien d'autre du Drive.
 */
const VERSION_API = '5.15';

// Vide : le script est dans le fichier (Extensions › Apps Script). Rempli seulement pour un script séparé.
const ID_FICHIER = '1TFXRVMLuVv_NauUH7VyJ4Pgd-Us5JuD3_3lzN_ojfvQ'; // fichier budget de Devlin

// Menu « Budget » dans le fichier (installation d'une nouvelle personne, 24/11/2026) : choisir sa clé sans toucher au code.
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('Budget').addItem('Choisir ma clé (installation)', 'choisirCle').addToUi(); } catch (e) {}
}
function choisirCle() {
  const ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties(), deja = p.getProperty('APP_KEY');
  const r = ui.prompt('Ma clé', (deja ? "Une clé existe déjà : la nouvelle la remplacera (l'appli te la redemandera).\n\n" : '') +
    "Choisis ta clé (au moins 6 caractères) : c'est le mot que tu taperas dans l'appli pour l'ouvrir.", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const cle = String(r.getResponseText() || '').trim();
  if (cle.length < 6) { ui.alert('Clé trop courte : au moins 6 caractères. Recommence par le menu Budget.'); return; }
  p.setProperty('APP_KEY', cle);
  ui.alert('Clé enregistrée ✓\n\nÉtape suivante : Extensions › Apps Script › Déployer › Nouveau déploiement.');
}

// À lancer depuis l'éditeur (choisis « verifier » en haut, puis ▶ Exécuter) :
// donne à Google toutes les autorisations dont le script a besoin et vérifie que tout répond.
function verifier() {
  const ss = classeur_();
  ['Paramètres', 'Journal', 'Échéancier', 'À venir'].forEach(n => { if (!ss.getSheetByName(n)) throw new Error('Onglet introuvable : ' + n); });
  CacheService.getScriptCache().put('test', '1', 60);
  LockService.getScriptLock().waitLock(5000); LockService.getScriptLock().releaseLock();
  const cle = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  const d = lireDonnees_();
  Logger.log('✓ Script ' + VERSION_API + ' relié à « ' + ss.getName() + ' » · ' + d.categories.length + ' catégories · ' + d.journal.length + ' lignes de Journal · clé ' + (cle ? 'enregistrée ✓' : 'MANQUANTE : menu Budget › Choisir ma clé'));
}

/* ---------- Points d'entrée ---------- */

// Ouvrir le lien /exec dans un navigateur affiche juste un message de contrôle.
function doGet() {
  return ContentService.createTextOutput('Budget · script ' + VERSION_API + ' en ligne ✓ (relié au fichier « ' + classeur_().getName() + ' »)');
}

function doPost(e) {
  let out, action = '?'; const t0 = Date.now();
  try {
    const req = JSON.parse(e.postData.contents); action = req && req.action;
    out = traiter_(req);
  } catch (err) {
    out = { ok: false, error: 'SERVEUR', message: String(err && err.message || err) };
    console.error('Erreur ' + action + ' : ' + out.message + (err && err.stack ? '\n' + err.stack : ''));
  }
  // Un lot : une ligne avec le nombre d'envois et d'échecs (le détail de chaque échec va dans « Détail »).
  const nom = action === 'lot' && out && out.ok && out.detail ? 'lot (' + out.detail.ops + ' envois, ' + out.detail.echecs + ' échec' + (out.detail.echecs > 1 ? 's' : '') + ')' : action;
  try { if (action !== 'data') console.log(nom + ' → ' + (out && out.ok ? 'ok' : 'refusé : ' + (out && out.message))); } catch (x) {}
  diagnostic_(nom, Date.now() - t0, out);
  if (out && out.detail) delete out.detail;
  let txt; try { txt = JSON.stringify(out); } catch (x) { txt = JSON.stringify({ ok: false, error: 'SERVEUR', message: 'Réponse illisible' }); }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function traiter_(req) {
  const cle = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (!cle) return { ok: false, error: 'SANS_CLE', message: 'Clé non configurée : dans ton fichier, menu Budget › Choisir ma clé.' };
  if (!req || req.key !== cle) return { ok: false, error: 'CLE', message: 'Clé incorrecte.' };

  if (req.action === 'data') return { ok: true, data: lireDonnees_() };
  if (req.action === 'lot') return avecVerrou_(() => lot_(req.ops));
  const f = ECRITURES_[req.action];
  if (!f) return { ok: false, error: 'ACTION', message: 'Action inconnue : ' + req.action };
  return avecVerrou_(() => f(req));
}

// Écritures : le corps seul, sans verrou (pris par traiter_ pour un envoi, une seule fois pour un lot).
const ECRITURES_ = {
  'add': r => ajouter_(r.ligne),
  'update': r => modifier_(r.row, r.attendu, r.ligne),
  'delete': r => supprimer_(r.row, r.attendu, r.uid),
  'ech-save': r => echEnregistrer_(r.row, r.attendu, r.e, r.uid),
  'ech-del': r => echSupprimer_(r.row, r.attendu),
  'budget': r => budget_(r.nom, r.budget, r.suivi),
  'cat-add': r => catAjouter_(r.nom, r.groupe, r.budget, r.suivi),
  'cat-edit': r => catModifier_(r.ancien, r.nom, r.groupe),
  'cat-del': r => catRetirer_(r.nom, r.cible),
  'solde-ctrl': r => soldeControle_(r.mois, r.solde, r.especes),
  'solde-dep': r => soldeDepart_(r.annee, r.valeur),
  'av-save': r => avEnregistrer_(r.row, r.attendu, r.a, r.uid),
  'av-del': r => avSupprimer_(r.row, r.attendu),
  'defi': r => defi_(r.mois, r.cat, r.montant),
  'pref-set': r => prefSet_(r.cle, r.valeur),
  'av-paye': r => avPaye_(r.row, r.attendu, r.ligne)
};

// v5.11 : plusieurs envois en une demande, exécutés dans l'ordre sous un seul verrou.
// Chaque envoi a sa propre réponse ; un envoi refusé ou en erreur n'empêche pas les suivants.
function lot_(ops) {
  if (!Array.isArray(ops) || !ops.length || ops.length > 50) return { ok: false, error: 'INVALIDE', message: 'Lot vide ou trop long.' };
  const erreurs = [];
  const res = ops.map((op, i) => {
    const f = op && ECRITURES_[op.action];
    let r;
    if (!f) r = { ok: false, error: 'ACTION', message: 'Action inconnue : ' + (op && op.action) };
    else {
      try { r = f(op) || { ok: false, error: 'SERVEUR', message: 'Réponse vide' }; }
      catch (err) { r = { ok: false, error: 'SERVEUR', message: String(err && err.message || err) }; console.error('Erreur lot ' + (op && op.action) + ' : ' + r.message); }
    }
    if (!r.ok) erreurs.push({ n: i + 1, action: op && op.action, error: r.error, message: r.message, detail: r.detail });
    if (r.detail) delete r.detail;
    return r;
  });
  return { ok: true, res, detail: { ops: ops.length, echecs: erreurs.length, actions: ops.map(o => o && o.action).join(','), erreurs } };
}

/* ---------- Lecture ---------- */

// Le classeur n'est ouvert qu'une fois par demande (avant : 2 ou 3 fois par ajout).
let SS_ = null;
function classeur_() {
  // D'abord le fichier qui contient le script (compatible avec @OnlyCurrentDoc) ; l'identifiant ne sert qu'à un script séparé.
  return SS_ || (SS_ = SpreadsheetApp.getActiveSpreadsheet() || (ID_FICHIER ? SpreadsheetApp.openById(ID_FICHIER) : null));
}

// Onglet caché « Diagnostic » : une ligne par demande (heure, action, durée, résultat), 300 dernières gardées.
// Permet à Claude de comprendre une lenteur ou un échec en lisant le fichier.
function diagnostic_(action, ms, out) {
  try {
    const ss = classeur_();
    let sh = ss.getSheetByName('Diagnostic');
    if (!sh) { sh = ss.insertSheet('Diagnostic'); sh.appendRow(['Heure', 'Action', 'Durée (ms)', 'Résultat', 'Message', 'Détail']); sh.hideSheet(); }
    const ok = out && out.ok;
    sh.appendRow([new Date(), String(action), ms, ok ? 'ok' : String(out && out.error || '?'), ok ? '' : String(out && out.message || '').slice(0, 200), out && out.detail ? JSON.stringify(out.detail).slice(0, 400) : '']);
    if (sh.getLastRow() > 400) sh.deleteRows(2, sh.getLastRow() - 301);
  } catch (e) { /* le diagnostic ne doit jamais bloquer un envoi */ }
}

function lireDonnees_() {
  const ss = classeur_();
  const par = ss.getSheetByName('Paramètres');
  const ech = ss.getSheetByName('Échéancier');
  const jou = ss.getSheetByName('Journal');

  const categories = par.getRange('A10:E40').getValues()
    .map((r, i) => ({
      ligne: 10 + i,
      nom: String(r[0]).trim(),
      groupe: String(r[1]).trim(),
      type: String(r[2]).trim(),
      budget: typeof r[3] === 'number' ? r[3] : null,
      suivi: String(r[4]).trim()
    }))
    .filter(c => c.nom && !/libre/i.test(c.nom));

  const mois = moisListe_(par).filter(String);

  const soldes = {};
  par.getRange('A52:B90').getValues().forEach(r => {
    if (r[0] !== '' && typeof r[1] === 'number') soldes[String(r[0])] = r[1];
  });

  let debut = null;
  const nr = ss.getRangeByName('DebutEcheancier');
  if (nr && nr.getValue() instanceof Date) debut = Utilities.formatDate(nr.getValue(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

  const echeancier = ech.getRange(5, 1, finEch_(ech) - 4, 10).getValues()
    .map((r, i) => ({
      ligne: 5 + i,
      lib: String(r[0]).trim(),
      cat: String(r[1]).trim(),
      montant: typeof r[2] === 'number' ? r[2] : null,
      freq: String(r[3]).trim(),
      jour: parseInt(String(r[4]).replace(/\D/g, ''), 10) || null,
      f: num_(r[5]), g: num_(r[6]), h: num_(r[7]), i: num_(r[8]),
      pause: r[9] === true
    }))
    .filter(e => e.lib);

  const nj = nbLignesJournal_(jou);
  const brut = jou.getRange(5, 1, nj, 4).getValues();
  const dates = jou.getMaxColumns() >= COL_DATE ? jou.getRange(5, COL_DATE, nj, 1).getValues() : [];
  const cp = colPaye_(jou, false);
  const payes = cp ? jou.getRange(5, cp, nj, 1).getValues() : [];
  const tz = ss.getSpreadsheetTimeZone();
  const journal = [];
  brut.forEach((r, i) => {
    if (r[0] === '' && r[1] === '' && r[2] === '' && r[3] === '') return;
    const d = dates[i] && dates[i][0];
    const l = { row: 5 + i, mois: String(r[0]), lib: String(r[1]), cat: String(r[2]), montant: r[3], date: d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : null };
    if (payes[i] && estEspeces_(payes[i][0])) l.especes = true; // champ ajouté seulement si vrai (réponse compacte)
    else if (payes[i] && estEnLigne_(payes[i][0])) l.enLigne = true;
    journal.push(l);
  });

  const libres = par.getRange('A10:A40').getValues().filter(r => !String(r[0]).trim() || /libre/i.test(String(r[0]))).length;
  const c = par.getRange('B43:B45').getValues();
  const controle = { mois: String(c[0][0]), solde: typeof c[1][0] === 'number' ? c[1][0] : null, especes: typeof c[2][0] === 'number' ? c[2][0] : null };

  const avenir = [];
  const shAv = ss.getSheetByName('À venir');
  if (shAv) shAv.getRange(4, 1, finAv_(shAv) - 3, 6).getValues().forEach((r, i) => {
    if (String(r[1]).trim() === '' && r[3] === '') return;
    avenir.push({ row: 4 + i, mois: String(r[0]), lib: String(r[1]).trim(), cat: String(r[2]).trim(), montant: typeof r[3] === 'number' ? r[3] : Number(r[3]) || 0, fait: r[4] === true, jour: num_(r[5]) });
  });

  return {
    categories, mois, soldes, debut, echeancier, journal, libres, controle, avenir, defis: defis_(), prefs: prefs_(), url: ss.getUrl(), apiVersion: VERSION_API,
    anomalies: jou.getRange('A3').getDisplayValue(),
    lu: new Date().toISOString()
  };
}

function num_(v) { return typeof v === 'number' ? v : (v === '' || v === null ? null : (Number(v) || null)); }

// Nombre de lignes du Journal à partir de la ligne 5 (la feuille rétrécit à chaque ligne supprimée).
function nbLignesJournal_(jou) { return Math.max(1, jou.getMaxRows() - 4); }
// Échéancier : lignes 5 → dernière ligne de l'onglet
function finEch_(ech) { return Math.max(5, ech.getMaxRows()); }
// À venir : lignes 4 → dernière ligne de l'onglet
function finAv_(sh) { return Math.max(4, sh.getMaxRows()); }
// Liste des mois du Journal (Paramètres, colonne N, de la ligne 4 jusqu'en bas)
function moisListe_(par) { return par.getRange(4, 14, Math.max(1, par.getMaxRows() - 3), 1).getDisplayValues().map(r => r[0]); }

/* ---------- Date de saisie (colonne M du Journal) ---------- */

const COL_DATE = 13;
function ecrireDate_(jou, row, iso) {
  if (jou.getMaxColumns() < COL_DATE) jou.insertColumnsAfter(jou.getMaxColumns(), COL_DATE - jou.getMaxColumns());
  const t = jou.getRange(4, COL_DATE);
  if (!t.getValue()) t.setValue('Saisi le');
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  const d = m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date();
  jou.getRange(row, COL_DATE).setValue(d).setNumberFormat('dd/MM/yyyy');
}

/* ---------- « Payé en » : dépense payée en espèces (colonne N du Journal, repérée par son en-tête) ---------- */

// La colonne est retrouvée par son titre « Payé en » en ligne 4 (N4 normalement).
// Si N4 contient déjà autre chose, on prend la première colonne après M dont l'en-tête ET les lignes sont vides :
// rien d'existant n'est jamais écrasé. Les colonnes A-D (contrôle d'identité) et M (date) ne changent pas.
const COL_PAYE = 14, TITRE_PAYE = 'Payé en';
function colPaye_(jou, creer) {
  const max = jou.getMaxColumns();
  if (max >= COL_PAYE) {
    const tetes = jou.getRange(4, COL_PAYE, 1, max - COL_PAYE + 1).getDisplayValues()[0].map(v => String(v).trim());
    const i = tetes.indexOf(TITRE_PAYE);
    if (i >= 0) return COL_PAYE + i;
    if (!creer) return 0;
    const bloc = jou.getRange(4, COL_PAYE, jou.getMaxRows() - 3, tetes.length).getDisplayValues();
    for (let k = 0; k < tetes.length; k++) {
      if (tetes[k] === '' && bloc.every(r => String(r[k]).trim() === '')) { jou.getRange(4, COL_PAYE + k).setValue(TITRE_PAYE); return COL_PAYE + k; }
    }
  }
  if (!creer) return 0;
  const c = Math.max(max + 1, COL_PAYE);
  jou.insertColumnsAfter(max, c - max);
  jou.getRange(4, c).setValue(TITRE_PAYE);
  return c;
}
function estEspeces_(v) { return /^esp[eè]ces$/i.test(String(v).trim()); }
function estEnLigne_(v) { return /^en\s*ligne$/i.test(String(v).trim()); }
// Écrit « Espèces », « En ligne » ou vide (les deux sont exclusifs : espèces l'emporte).
// Une ligne ni en espèces ni en ligne ne crée jamais la colonne.
function ecrirePaye_(jou, row, especes, enLigne) {
  const v = especes ? 'Espèces' : (enLigne ? 'En ligne' : '');
  const c = colPaye_(jou, !!v);
  if (c) jou.getRange(row, c).setValue(v);
}
// Modification : une ancienne appli qui n'envoie que « especes » ne doit pas effacer un « En ligne ».
function majPaye_(jou, row, l) {
  const e = typeof l.especes === 'boolean', n = typeof l.enLigne === 'boolean';
  if (!e && !n) return;
  if (!n) {
    const c = colPaye_(jou, false);
    if (!l.especes && c && estEnLigne_(jou.getRange(row, c).getValue())) return;
  }
  ecrirePaye_(jou, row, l.especes === true, l.enLigne === true);
}

/* ---------- Préférences de l'appli (tirelires, voyage, escalier, séries…) gardées dans le script ---------- */

// « especes » : { depuis, base, deja, r: [{ d, m }] } (suivi des espèces entre deux contrôles du solde)
// « envies » : { l: [{ i, l, p, c, d, s: 'att'|'ach'|'ren', f }], ren: { n, eco }, ach: { n, eco }, r: { taux, seuil, delai, env } }
//              (liste d'envies avec délai, anciennes envies résumées en compteurs, réglages « Achats plaisir »)
const PREFS_OK_ = ['tirelires', 'voyages', 'escalier', 'semaines', 'objectifs', 'jeu', 'especes', 'envies'];
function prefs_() {
  const p = PropertiesService.getScriptProperties(), o = {};
  PREFS_OK_.forEach(k => { try { const v = p.getProperty('PREF_' + k); if (v) o[k] = JSON.parse(v); } catch (e) {} });
  return o;
}
function prefSet_(cle, valeur) {
  if (PREFS_OK_.indexOf(String(cle)) < 0) return { ok: false, error: 'INVALIDE', message: 'Préférence inconnue.' };
  const t = JSON.stringify(valeur == null ? null : valeur);
  if (t.length > 8000) return { ok: false, error: 'INVALIDE', message: 'Trop de données.' };
  const p = PropertiesService.getScriptProperties();
  if (valeur == null) p.deleteProperty('PREF_' + cle); else p.setProperty('PREF_' + cle, t);
  return { ok: true };
}

/* ---------- Défis du mois (gardés dans les propriétés du script) ---------- */

function defis_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('DEFIS') || '{}'); } catch (e) { return {}; }
}
function defi_(mois, cat, montant) {
  if (!/^\d{4}-\d{1,2}$/.test(String(mois))) return { ok: false, error: 'INVALIDE', message: 'Mois du défi incorrect.' };
  const d = defis_();
  if (!cat) delete d[mois];
  else {
    const v = Math.round(Number(montant) * 100) / 100;
    if (!(v >= 0)) return { ok: false, error: 'INVALIDE', message: 'Montant du défi incorrect.' };
    d[mois] = { cat: String(cat), montant: v };
  }
  PropertiesService.getScriptProperties().setProperty('DEFIS', JSON.stringify(d));
  return { ok: true };
}

/* ---------- Écriture ---------- */

// v5.14 · Mémoire durable des envois déjà faits (uid → ligne), pour qu'un envoi rejoué après une coupure
// (le fichier a écrit, la réponse s'est perdue) ne soit pas appliqué deux fois, même le lendemain.
// Propriété « UIDS » : { "type:uid": [ligne, decale(0/1), jour] }, 30 jours et ~8 000 caractères au plus (les plus anciens partent).
// Toujours lue et écrite sous le verrou (avecVerrou_) : retenue en mémoire pendant la demande, écrite une fois à la fin.
const UIDS_JOURS = 30, UIDS_MAX = 8000;
let UIDS_ = null, UIDS_SALE_ = false;
function uidsLire_() {
  if (UIDS_) return UIDS_;
  try { UIDS_ = JSON.parse(PropertiesService.getScriptProperties().getProperty('UIDS') || '{}') || {}; } catch (e) { UIDS_ = {}; }
  return UIDS_;
}
function cleUid_(type, uid) { return uid ? type + ':' + String(uid).slice(0, 40) : null; }
function uidVu_(cle) { return cle ? uidsLire_()[cle] || null : null; }
function uidRetenir_(cle, row, decale) {
  if (!cle) return;
  const m = uidsLire_(), j = Math.floor(Date.now() / 86400000);
  m[cle] = [row || 0, decale ? 1 : 0, j];
  UIDS_SALE_ = true; // écrit une seule fois, à la fin de la demande (uidsEcrire_, sous le verrou), même pour un lot de 20 envois
}
function uidsEcrire_() {
  if (!UIDS_SALE_ || !UIDS_) return;
  UIDS_SALE_ = false;
  const m = UIDS_, j = Math.floor(Date.now() / 86400000);
  let cles = Object.keys(m).filter(k => !(Array.isArray(m[k]) && j - m[k][2] <= UIDS_JOURS));
  cles.forEach(k => delete m[k]);
  let t = JSON.stringify(m);
  if (t.length > UIDS_MAX) {
    cles = Object.keys(m).sort((a, b) => m[a][2] - m[b][2]);
    while (t.length > UIDS_MAX && cles.length > 1) { delete m[cles.shift()]; t = JSON.stringify(m); }
  }
  PropertiesService.getScriptProperties().setProperty('UIDS', t);
}

// v5.14 · Libellé écrit en texte : sinon Sheets transforme « =x », « +3 », « 1,20 », « 12/11 », « VRAI »…
// en formule, nombre ou date, et la ligne ne se reconnaît plus ensuite (« La ligne a changé »).
// La cellule du libellé passe au format « Texte brut » AVANT d'écrire, puis le texte est écrit tel quel
// (pas d'apostrophe devant : dans une cellule déjà en texte brut, elle resterait collée au libellé).
function libEnTexte_(sh, row, col) { sh.getRange(row, col).setNumberFormat('@'); }

function avecVerrou_(fn) {
  const lock = LockService.getScriptLock();
  // Fichier occupé par un autre envoi : on le dit à l'appli, qui réessaiera toute seule (rien n'est perdu).
  if (!lock.tryLock(20000)) return { ok: false, error: 'OCCUPE', message: 'Fichier occupé, nouvel essai automatique.' };
  try { return fn(); } finally {
    try { uidsEcrire_(); } catch (err) { console.error('UIDS non enregistrés : ' + (err && err.message || err)); }
    lock.releaseLock();
  }
}

function valider_(l, futurOk) {
  const ss = classeur_();
  const par = ss.getSheetByName('Paramètres');
  const mois = moisListe_(par);
  const cats = par.getRange('A10:A40').getValues().map(r => String(r[0]).trim());
  if (!l) throw new Error('Ligne vide.');
  // « À venir » : un mois futur pas encore dans la liste du Journal est accepté s'il est bien écrit (v5.10).
  const moisBienEcrit = /^(Janvier|Février|Mars|Avril|Mai|Juin|Juillet|Août|Septembre|Octobre|Novembre|Décembre) \d{4}$/.test(String(l.mois));
  if (mois.indexOf(l.mois) < 0 && !(futurOk && moisBienEcrit)) return 'Mois inconnu : ' + l.mois;
  if (cats.indexOf(l.cat) < 0) return 'Catégorie inconnue : ' + l.cat;
  const m = Number(l.montant);
  const ecart = /^Écart/.test(String(l.lib || ''));
  const avance = /^Avances/i.test(String(l.cat || ''));
  if (!(m > 0) && !((ecart || avance) && m < 0)) return 'Montant incorrect.';
  return null;
}

function identique_(vals, a) {
  return a && String(vals[0]) === String(a.mois) && String(vals[1]) === String(a.lib) &&
    String(vals[2]) === String(a.cat) && Math.abs(Number(vals[3]) - Number(a.montant)) < 0.005;
}

function ajouter_(l) {
  const err = valider_(l);
  if (err) return { ok: false, error: 'INVALIDE', message: err };
  // Anti-doublon : si le téléphone renvoie un ajout déjà écrit (réseau coupé avant la réponse), on ne l'écrit pas deux fois.
  // Mémoire durable (UIDS, 30 jours).
  const jou = classeur_().getSheetByName('Journal');
  const cleUid = cleUid_('add', l.uid), vu = uidVu_(cleUid);
  if (vu) return { ok: true, row: ligneRetrouvee_(jou, vu[0], l), deja: true };
  const vals = jou.getRange(5, 1, nbLignesJournal_(jou), 4).getValues();
  let dernier = -1;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i].some(v => v !== '')) { dernier = i; break; }
  }
  const row = 5 + dernier + 1;
  if (row > jou.getMaxRows()) jou.insertRowsAfter(jou.getMaxRows(), 100);
  libEnTexte_(jou, row, 2);
  jou.getRange(row, 1, 1, 4).setValues([[l.mois, String(l.lib || l.cat), l.cat, Math.round(Number(l.montant) * 100) / 100]]);
  ecrireDate_(jou, row, l.date);
  ecrirePaye_(jou, row, l.especes === true, l.enLigne === true);
  uidRetenir_(cleUid, row, false);
  return { ok: true, row };
}

// Un ajout déjà fait : sa ligne a pu remonter depuis (suppressions au-dessus). On cherche la ligne identique
// la plus proche, à son numéro ou au-dessus ; introuvable (supprimée ou changée entre-temps) : null, l'appli relira le fichier.
function ligneRetrouvee_(jou, row, l) {
  row = Number(row) || 0;
  const n = Math.min(row, jou.getMaxRows()) - 4;
  if (n < 1) return null;
  const vals = jou.getRange(5, 1, n, 4).getValues();
  for (let i = n - 1; i >= 0; i--) if (identique_(vals[i], l)) return 5 + i;
  return null;
}

function modifier_(row, attendu, l) {
  const err = valider_(l);
  if (err) return { ok: false, error: 'INVALIDE', message: err };
  const jou = classeur_().getSheetByName('Journal');
  if (!(Number(row) >= 5)) return { ok: false, error: 'INVALIDE', message: 'Ligne inconnue.' };
  const r = jou.getRange(row, 1, 1, 4), v = r.getValues()[0];
  // v5.14 : une modification déjà faite (renvoyée après une coupure) n'est plus refusée : on la réécrit à l'identique.
  if (!identique_(v, attendu) && !identique_(v, Object.assign({}, l, { lib: l.lib || l.cat }))) return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans le fichier.', detail: { row, fichier: v.map(String), attendu } };
  libEnTexte_(jou, row, 2);
  r.setValues([[l.mois, String(l.lib || l.cat), l.cat, Math.round(Number(l.montant) * 100) / 100]]);
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(l.date || ''))) ecrireDate_(jou, row, l.date); // date de la dépense corrigée dans l'appli
  majPaye_(jou, row, l); // une ancienne appli qui n'envoie rien ne touche pas à la colonne
  return { ok: true, row };
}

function supprimer_(row, attendu, uid) {
  // v5.14 : suppression déjà faite (renvoyée après une coupure) : on ne supprime pas la ligne suivante identique.
  const cleUid = cleUid_('del', uid), vu = uidVu_(cleUid);
  if (vu) return { ok: true, row: vu[0], decale: vu[1] === 1, deja: true };
  if (!(Number(row) >= 5)) return { ok: false, error: 'INVALIDE', message: 'Ligne inconnue.' };
  const jou = classeur_().getSheetByName('Journal');
  const r = jou.getRange(row, 1, 1, 4);
  if (!identique_(r.getValues()[0], attendu)) return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans le fichier.', detail: { row, fichier: r.getValues()[0].map(String), attendu } };
  if (row > 5) jou.deleteRow(row); else {
    r.clearContent();
    if (jou.getMaxColumns() >= COL_DATE) jou.getRange(row, COL_DATE).clearContent();
    const cp = colPaye_(jou, false); if (cp) jou.getRange(row, cp).clearContent();
  }
  uidRetenir_(cleUid, row, row > 5);
  return { ok: true, row, decale: row > 5 };
}

/* ---------- Échéancier ---------- */

const FREQS_ = ['Mensuel', 'Trimestriel', 'Semestriel', 'Annuel'];

function echIdentique_(v, a) {
  return a && String(v[0]).trim() === String(a.lib) && String(v[1]).trim() === String(a.cat) &&
    Math.abs(Number(v[2]) - Number(a.montant)) < 0.005;
}

function echEnregistrer_(row, attendu, e, uid) {
  const ss = classeur_(), ech = ss.getSheetByName('Échéancier');
  const cats = ss.getSheetByName('Paramètres').getRange('A10:A40').getValues().map(r => String(r[0]).trim());
  if (!e || !String(e.lib || '').trim()) return { ok: false, error: 'INVALIDE', message: 'Libellé manquant.' };
  if (cats.indexOf(e.cat) < 0) return { ok: false, error: 'INVALIDE', message: 'Catégorie inconnue : ' + e.cat };
  if (!(Number(e.montant) > 0)) return { ok: false, error: 'INVALIDE', message: 'Montant incorrect.' };
  if (FREQS_.indexOf(e.freq) < 0) return { ok: false, error: 'INVALIDE', message: 'Fréquence inconnue.' };
  const mois = v => (v === null || v === '' || v === undefined) ? '' : Math.min(12, Math.max(1, Math.round(Number(v))));
  const an = v => (v === null || v === '' || v === undefined) ? '' : Math.round(Number(v));
  const ligne = [String(e.lib).trim(), e.cat, Math.round(Number(e.montant) * 100) / 100, e.freq,
    e.jour ? Math.min(31, Math.max(1, Math.round(Number(e.jour)))) : '',
    mois(e.f) || 1, mois(e.g) || 12, an(e.h), an(e.i), e.pause ? true : ''];
  // v5.14 : nouvelle ligne renvoyée après une coupure (déjà écrite) : on renvoie sa ligne au lieu d'en créer une 2e.
  const cleUid = row ? null : cleUid_('ech', uid), vu = uidVu_(cleUid);
  if (vu) return { ok: true, row: vu[0], deja: true };
  if (!row) {
    const a = ech.getRange(5, 1, finEch_(ech) - 4, 1).getValues();
    let i = a.findIndex(r => String(r[0]).trim() === '');
    if (i < 0) return { ok: false, error: 'PLEIN', message: "L'Échéancier est plein (" + a.length + " lignes). Supprime une ligne terminée." };
    row = 5 + i;
  } else {
    if (row < 5 || row > finEch_(ech)) return { ok: false, error: 'INVALIDE', message: 'Ligne hors Échéancier.' };
    const v = ech.getRange(row, 1, 1, 3).getValues()[0];
    if (!echIdentique_(v, attendu) && !echIdentique_(v, { lib: String(e.lib).trim(), cat: e.cat, montant: e.montant })) return { ok: false, error: 'CHANGE', message: "La ligne a changé dans l'Échéancier." };
  }
  libEnTexte_(ech, row, 1);
  ech.getRange(row, 1, 1, 10).setValues([ligne]);
  uidRetenir_(cleUid, row, false);
  return { ok: true, row };
}

function echSupprimer_(row, attendu) {
  const ech = classeur_().getSheetByName('Échéancier');
  if (row < 5 || row > finEch_(ech)) return { ok: false, error: 'INVALIDE', message: 'Ligne hors Échéancier.' };
  const v = ech.getRange(row, 1, 1, 3).getValues()[0];
  if (v.every(x => String(x).trim() === '')) return { ok: true, row, deja: true }; // v5.14 : déjà supprimée (envoi renvoyé)
  if (!echIdentique_(v, attendu)) return { ok: false, error: 'CHANGE', message: "La ligne a changé dans l'Échéancier." };
  ech.getRange(row, 1, 1, 10).clearContent();
  return { ok: true, row };
}

/* ---------- Budgets et catégories ---------- */

const GROUPES_ = ['Revenus', 'Charges fixes', 'Abonnements', 'Vie courante', 'Loisirs', 'Divers', 'Épargne', 'Avances'];
const SUIVIS_ = ['Mensuel', 'Annuel', 'Sans alerte'];

function ligneCat_(par, nom) {
  const a = par.getRange('A10:A40').getValues().map(r => String(r[0]).trim());
  const i = a.indexOf(String(nom).trim());
  return i < 0 ? -1 : 10 + i;
}

function budget_(nom, budget, suivi) {
  const par = classeur_().getSheetByName('Paramètres');
  const r = ligneCat_(par, nom);
  if (r < 0) return { ok: false, error: 'INVALIDE', message: 'Catégorie inconnue : ' + nom };
  if (SUIVIS_.indexOf(suivi) < 0) return { ok: false, error: 'INVALIDE', message: 'Suivi inconnu.' };
  const b = Number(budget);
  par.getRange(r, 4, 1, 2).setValues([[b > 0 ? Math.round(b * 100) / 100 : '', suivi]]);
  return { ok: true };
}

function catAjouter_(nom, groupe, budget, suivi) {
  nom = String(nom || '').trim();
  const par = classeur_().getSheetByName('Paramètres');
  if (!nom) return { ok: false, error: 'INVALIDE', message: 'Nom vide.' };
  if (GROUPES_.indexOf(groupe) < 0) return { ok: false, error: 'INVALIDE', message: 'Groupe inconnu.' };
  if (ligneCat_(par, nom) > 0) return { ok: false, error: 'INVALIDE', message: 'Cette catégorie existe déjà.' };
  const a = par.getRange('A10:A40').getValues();
  const i = a.findIndex(r => !String(r[0]).trim() || /libre/i.test(String(r[0])));
  if (i < 0) return { ok: false, error: 'PLEIN', message: 'Plus de place : retire une catégorie pour en libérer une.' };
  const r = 10 + i, b = Number(budget);
  par.getRange(r, 1, 1, 2).setValues([[nom, groupe]]);
  par.getRange(r, 4, 1, 2).setValues([[b > 0 ? b : '', SUIVIS_.indexOf(suivi) >= 0 ? suivi : 'Mensuel']]);
  return { ok: true, row: r };
}

// Remplace un nom de catégorie partout (Journal, Échéancier, À venir).
function remplacerCat_(ancien, nouveau) {
  const ss = classeur_(); let n = 0;
  [['Journal', 'C5:C' + ss.getSheetByName('Journal').getMaxRows()], ['Échéancier', 'B5:B'], ['À venir', 'C4:C']].forEach(([f, plage]) => {
    const sh = ss.getSheetByName(f); if (!sh) return;
    if (/:[A-Z]$/.test(plage)) plage += sh.getMaxRows();
    const rg = sh.getRange(plage), v = rg.getValues(); let ch = false;
    v.forEach(r => { if (String(r[0]).trim() === ancien) { r[0] = nouveau; ch = true; n++; } });
    if (ch) rg.setValues(v);
  });
  // Le défi du mois suit aussi le nouveau nom (v5.9).
  const d = defis_(); let dch = false;
  Object.keys(d).forEach(k => { if (d[k] && d[k].cat === ancien) { d[k].cat = nouveau; dch = true; } });
  if (dch) PropertiesService.getScriptProperties().setProperty('DEFIS', JSON.stringify(d));
  return n;
}

function catModifier_(ancien, nom, groupe) {
  nom = String(nom || '').trim();
  const par = classeur_().getSheetByName('Paramètres');
  const r = ligneCat_(par, ancien);
  if (r < 0) return { ok: false, error: 'INVALIDE', message: 'Catégorie inconnue : ' + ancien };
  if (!nom) return { ok: false, error: 'INVALIDE', message: 'Nom vide.' };
  if (nom !== ancien && ligneCat_(par, nom) > 0) return { ok: false, error: 'INVALIDE', message: 'Ce nom existe déjà.' };
  if (groupe && GROUPES_.indexOf(groupe) < 0) return { ok: false, error: 'INVALIDE', message: 'Groupe inconnu.' };
  par.getRange(r, 1, 1, 2).setValues([[nom, groupe || par.getRange(r, 2).getValue()]]);
  const n = nom !== ancien ? remplacerCat_(ancien, nom) : 0;
  return { ok: true, lignes: n };
}

function catRetirer_(nom, cible) {
  const par = classeur_().getSheetByName('Paramètres');
  const r = ligneCat_(par, nom);
  if (r < 0) return { ok: false, error: 'INVALIDE', message: 'Catégorie inconnue : ' + nom };
  if (cible) {
    if (ligneCat_(par, cible) < 0 || cible === nom) return { ok: false, error: 'INVALIDE', message: 'Catégorie de destination inconnue.' };
    remplacerCat_(nom, cible);
  } else {
    const ss = classeur_();
    const usage = ss.getSheetByName('Journal').getRange(5, 3, nbLignesJournal_(ss.getSheetByName('Journal')), 1).getValues().filter(x => String(x[0]).trim() === nom).length +
      (sh => sh.getRange(5, 2, finEch_(sh) - 4, 1).getValues())(ss.getSheetByName('Échéancier')).filter(x => String(x[0]).trim() === nom).length +
      (sh => sh ? sh.getRange(4, 3, finAv_(sh) - 3, 3).getValues().filter(x => String(x[0]).trim() === nom && x[2] !== true).length : 0)(ss.getSheetByName('À venir'));
    if (usage) return { ok: false, error: 'UTILISEE', message: usage + ' ligne(s) utilisent encore cette catégorie.' };
  }
  par.getRange(r, 1, 1, 2).setValues([['Catégorie libre', '']]);
  par.getRange(r, 4, 1, 2).setValues([['', 'Annuel']]);
  return { ok: true };
}

/* ---------- Soldes ---------- */

function soldeControle_(mois, solde, especes) {
  const par = classeur_().getSheetByName('Paramètres');
  const liste = par.getRange('I4:I15').getValues().map(r => String(r[0]));
  if (liste.indexOf(mois) < 0) return { ok: false, error: 'INVALIDE', message: 'Mois inconnu : ' + mois };
  par.getRange('B43:B45').setValues([[mois], [solde === null || solde === '' ? '' : Number(solde)], [especes ? Number(especes) : '']]);
  return { ok: true };
}

function soldeDepart_(annee, valeur) {
  const par = classeur_().getSheetByName('Paramètres');
  const a = par.getRange('A52:A90').getValues();
  const i = a.findIndex(r => String(r[0]) === String(annee));
  if (i < 0) return { ok: false, error: 'INVALIDE', message: 'Année absente de Paramètres ③ : ' + annee };
  par.getRange(52 + i, 2).setValue(Number(valeur));
  return { ok: true };
}

/* ---------- Achats prévus (feuille « À venir ») ---------- */

function avIdentique_(v, a) {
  return a && String(v[1]).trim() === String(a.lib) && String(v[2]).trim() === String(a.cat) && Math.abs(Number(v[3]) - Number(a.montant)) < 0.005;
}

function avEnregistrer_(row, attendu, a, uid) {
  const ss = classeur_(), sh = ss.getSheetByName('À venir');
  if (!sh) return { ok: false, error: 'INVALIDE', message: 'Feuille « À venir » introuvable.' };
  const err = valider_({ mois: a.mois, lib: a.lib, cat: a.cat, montant: a.montant }, true);
  if (err) return { ok: false, error: 'INVALIDE', message: err };
  // v5.14 : achat prévu renvoyé après une coupure (déjà écrit) : pas de 2e ligne.
  const cleUid = row ? null : cleUid_('av', uid), vu = uidVu_(cleUid);
  if (vu) return { ok: true, row: vu[0], deja: true };
  if (!row) {
    const v = sh.getRange(4, 1, finAv_(sh) - 3, 4).getValues();
    const i = v.findIndex(r => String(r[1]).trim() === '' && r[3] === '');
    if (i < 0) return { ok: false, error: 'PLEIN', message: 'La feuille « À venir » est pleine.' };
    row = 4 + i;
  } else if (!avIdentique_(sh.getRange(row, 1, 1, 4).getValues()[0], attendu) && !avIdentique_(sh.getRange(row, 1, 1, 4).getValues()[0], { lib: String(a.lib).trim(), cat: a.cat, montant: a.montant })) {
    return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans « À venir ».' };
  }
  // Écrit cellule par cellule : en cas de souci (validation, format…), le message dit laquelle.
  const vals = [a.mois, String(a.lib).trim(), a.cat, Math.round(Number(a.montant) * 100) / 100];
  // Le mois est forcé en texte : sinon Sheets transforme « Novembre 2026 » en date et le menu le refuse.
  sh.getRange(row, 1).setNumberFormat('@');
  libEnTexte_(sh, row, 2);
  ['Mois', 'Libellé', 'Catégorie', 'Montant'].forEach((nom, k) => {
    try { sh.getRange(row, k + 1).setValue(vals[k]); } catch (err) { throw new Error('« À venir », colonne ' + nom + ' : ' + err.message); }
  });
  sh.getRange(row, 5).setValue(false);
  sh.getRange(row, 6).setValue(a.jour ? Math.round(Number(a.jour)) : '');
  uidRetenir_(cleUid, row, false);
  return { ok: true, row };
}

function avSupprimer_(row, attendu) {
  const sh = classeur_().getSheetByName('À venir');
  const v = sh.getRange(row, 1, 1, 4).getValues()[0];
  if (String(v[1]).trim() === '' && v[3] === '') return { ok: true, deja: true }; // v5.14 : déjà supprimé (envoi renvoyé)
  if (!avIdentique_(v, attendu)) return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans « À venir ».' };
  sh.getRange(row, 1, 1, 4).clearContent(); sh.getRange(row, 6).clearContent(); sh.getRange(row, 5).setValue(false);
  return { ok: true };
}

function avPaye_(row, attendu, ligne) {
  const sh = classeur_().getSheetByName('À venir');
  if (!avIdentique_(sh.getRange(row, 1, 1, 4).getValues()[0], attendu)) return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans « À venir ».' };
  const r = ajouter_(ligne);
  if (!r.ok) return r;
  sh.getRange(row, 5).setValue(true);
  return { ok: true, row: r.row };
}