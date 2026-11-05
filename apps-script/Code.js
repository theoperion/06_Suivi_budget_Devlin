/**
 * BUDGET · API du budget de Devlin (Apps Script) · v5.10
 * ------------------------------------------
 * Ce script est relié au fichier budget dans lequel il a été créé.
 * Il lit les Paramètres, l'Échéancier et le Journal, et écrit dans le Journal.
 * Toutes les demandes doivent contenir la clé d'accès.
 *
 * v5.0 : plus aucun lien avec le fichier « Portefeuille & Épargne ».
 *        Un retrait de livret se saisit comme un revenu (ex. « Autres revenus », libellé « Retrait livret A »).
 * v5.1 : anti-doublon sur les ajouts (« payé pour d'autres » qui s'enregistrait deux fois).
 * v5.2 : la date de la dépense (colonne « Saisi le ») se modifie aussi quand on corrige une ligne.
 * v5.3 : fonction verifier() pour donner les autorisations et tester le script en un clic.
 * v5.6 : « À venir » : le mois est écrit en texte (corrige l'échec des éléments à venir).
 * v5.10 : « À venir » accepte un mois futur pas encore présent dans la liste du Journal.
 * v5.9 : renommer une catégorie met aussi à jour le défi du mois ; on ne peut plus retirer une catégorie encore utilisée dans « À venir ».
 * v5.8 : classeur ouvert une seule fois par demande ; onglet caché « Diagnostic » (durée et résultat de chaque demande).
 * v5.7 : fichier occupé → réponse OCCUPE (l'appli réessaie) au lieu d'une erreur qui faisait perdre la saisie.
 * v5.5 : chaque appel est noté dans Apps Script › Exécutions (action, résultat, erreur) pour diagnostiquer.
 * v5.4 : la colonne « Contrôle » du Journal ignore les lignes validées « c'est voulu » (libellé terminé par ✓). Mise à jour automatique, une seule fois.
 *
 * Installation :
 *  1. Remplace CLE_A_ENREGISTRER par ta clé, lance initialiserCle() une fois (inutile si c'est déjà fait).
 *  2. Déployer > Gérer les déploiements > crayon > Nouvelle version > Déployer
 *     (Exécuter en tant que : moi · Accès : tout le monde).
 */

const CLE_A_ENREGISTRER = 'CHANGE-MOI';
const VERSION_API = '5.10';

// Laisse vide si le script est ouvert depuis le fichier (Extensions > Apps Script).
const ID_FICHIER = '1TFXRVMLuVv_NauUH7VyJ4Pgd-Us5JuD3_3lzN_ojfvQ'; // fichier budget de Devlin

function initialiserCle() {
  if (!CLE_A_ENREGISTRER || CLE_A_ENREGISTRER === 'CHANGE-MOI') {
    throw new Error('Remplace d\'abord CHANGE-MOI par ta clé (ligne 17).');
  }
  PropertiesService.getScriptProperties().setProperty('APP_KEY', CLE_A_ENREGISTRER);
  Logger.log('Clé enregistrée. Tu peux remettre CHANGE-MOI à la ligne 17.');
}

// À lancer depuis l'éditeur (choisis « verifier » en haut, puis ▶ Exécuter) :
// donne à Google toutes les autorisations dont le script a besoin et vérifie que tout répond.
function verifier() {
  const ss = classeur_();
  ['Paramètres', 'Journal', 'Échéancier', 'À venir'].forEach(n => { if (!ss.getSheetByName(n)) throw new Error('Onglet introuvable : ' + n); });
  CacheService.getScriptCache().put('test', '1', 60);
  PropertiesService.getScriptProperties().deleteProperty('FORMULE_L4_V'); majFormules_(ss);
  Logger.log('Formule Contrôle (L4) : ' + (ss.getSheetByName('Journal').getRange('L4').getFormula().indexOf('✓') >= 0 ? 'à jour ✓' : 'NON MODIFIÉE (préviens Claude)'));
  LockService.getScriptLock().waitLock(5000); LockService.getScriptLock().releaseLock();
  const cle = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  const d = lireDonnees_();
  Logger.log('✓ Script ' + VERSION_API + ' relié à « ' + ss.getName() + ' » · ' + d.categories.length + ' catégories · ' + d.journal.length + ' lignes de Journal · clé ' + (cle ? 'enregistrée ✓' : 'MANQUANTE : lance initialiserCle()'));
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
  try { if (action !== 'data') console.log(action + ' → ' + (out && out.ok ? 'ok' : 'refusé : ' + (out && out.message))); } catch (x) {}
  diagnostic_(action, Date.now() - t0, out);
  if (out && out.detail) delete out.detail;
  let txt; try { txt = JSON.stringify(out); } catch (x) { txt = JSON.stringify({ ok: false, error: 'SERVEUR', message: 'Réponse illisible' }); }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

// Appelée par l'interface quand elle tourne directement dans Apps Script.
function api(texte) {
  try {
    return JSON.stringify(traiter_(JSON.parse(texte)));
  } catch (err) {
    return JSON.stringify({ ok: false, error: 'SERVEUR', message: String(err && err.message || err) });
  }
}

function traiter_(req) {
  const cle = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (!cle) return { ok: false, error: 'SANS_CLE', message: 'Clé non configurée : lance initialiserCle() dans Apps Script.' };
  if (!req || req.key !== cle) return { ok: false, error: 'CLE', message: 'Clé incorrecte.' };

  switch (req.action) {
    case 'data': return { ok: true, data: lireDonnees_() };
    case 'add': return avecVerrou_(() => ajouter_(req.ligne));
    case 'update': return avecVerrou_(() => modifier_(req.row, req.attendu, req.ligne));
    case 'delete': return avecVerrou_(() => supprimer_(req.row, req.attendu));
    case 'ech-save': return avecVerrou_(() => echEnregistrer_(req.row, req.attendu, req.e));
    case 'ech-del': return avecVerrou_(() => echSupprimer_(req.row, req.attendu));
    case 'budget': return avecVerrou_(() => budget_(req.nom, req.budget, req.suivi));
    case 'cat-add': return avecVerrou_(() => catAjouter_(req.nom, req.groupe, req.budget, req.suivi));
    case 'cat-edit': return avecVerrou_(() => catModifier_(req.ancien, req.nom, req.groupe));
    case 'cat-del': return avecVerrou_(() => catRetirer_(req.nom, req.cible));
    case 'solde-ctrl': return avecVerrou_(() => soldeControle_(req.mois, req.solde, req.especes));
    case 'solde-dep': return avecVerrou_(() => soldeDepart_(req.annee, req.valeur));
    case 'av-save': return avecVerrou_(() => avEnregistrer_(req.row, req.attendu, req.a));
    case 'av-del': return avecVerrou_(() => avSupprimer_(req.row, req.attendu));
    case 'defi': return avecVerrou_(() => defi_(req.mois, req.cat, req.montant));
    case 'pref-set': return avecVerrou_(() => prefSet_(req.cle, req.valeur));
    case 'av-paye': return avecVerrou_(() => avPaye_(req.row, req.attendu, req.ligne));
    default: return { ok: false, error: 'ACTION', message: 'Action inconnue : ' + req.action };
  }
}

/* ---------- Lecture ---------- */

// Le classeur n'est ouvert qu'une fois par demande (avant : 2 ou 3 fois par ajout).
let SS_ = null;
function classeur_() {
  return SS_ || (SS_ = ID_FICHIER ? SpreadsheetApp.openById(ID_FICHIER) : SpreadsheetApp.getActiveSpreadsheet());
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
  majFormules_(ss);
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
  const tz = ss.getSpreadsheetTimeZone();
  const journal = [];
  brut.forEach((r, i) => {
    if (r[0] === '' && r[1] === '' && r[2] === '' && r[3] === '') return;
    const d = dates[i] && dates[i][0];
    journal.push({ row: 5 + i, mois: String(r[0]), lib: String(r[1]), cat: String(r[2]), montant: r[3], date: d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : null });
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

// Colonne « Contrôle » du Journal (L4) : une ligne dont le libellé finit par ✓ n'est plus signalée « déjà dans l'Échéancier ? ».
// Corrige la formule une seule fois, sans rien toucher d'autre.
function majFormules_(ss) {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('FORMULE_L4_V') === '1') return;
  try {
    const c = (ss || classeur_()).getSheetByName('Journal').getRange('L4'), f = c.getFormula();
    let ok = !f || f.indexOf('✓') >= 0;
    if (!ok) {
      const g = f.replace(/IF\(\(A5:A=""\)\+\(cnt=0\)\s*,/, 'IF((A5:A="")+(cnt=0)+REGEXMATCH(B5:B&"","✓\\s*$"),');
      if (g !== f) { c.setFormula(g); ok = true; }
    }
    if (ok) p.setProperty('FORMULE_L4_V', '1');
  } catch (e) { /* on réessaiera au prochain chargement */ }
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

/* ---------- Préférences de l'appli (tirelires, voyage, escalier, séries…) gardées dans le script ---------- */

const PREFS_OK_ = ['tirelires', 'voyages', 'escalier', 'semaines', 'objectifs', 'jeu'];
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

function avecVerrou_(fn) {
  const lock = LockService.getScriptLock();
  // Fichier occupé par un autre envoi : on le dit à l'appli, qui réessaiera toute seule (rien n'est perdu).
  if (!lock.tryLock(20000)) return { ok: false, error: 'OCCUPE', message: 'Fichier occupé, nouvel essai automatique.' };
  try { return fn(); } finally { lock.releaseLock(); }
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
  const cache = CacheService.getScriptCache(), cleUid = l.uid ? 'add_' + String(l.uid).slice(0, 40) : null;
  if (cleUid && cache.get(cleUid)) return { ok: true, row: Number(cache.get(cleUid)), deja: true };
  const jou = classeur_().getSheetByName('Journal');
  const vals = jou.getRange(5, 1, nbLignesJournal_(jou), 4).getValues();
  let dernier = -1;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i].some(v => v !== '')) { dernier = i; break; }
  }
  const row = 5 + dernier + 1;
  if (row > jou.getMaxRows()) jou.insertRowsAfter(jou.getMaxRows(), 100);
  jou.getRange(row, 1, 1, 4).setValues([[l.mois, l.lib || l.cat, l.cat, Math.round(Number(l.montant) * 100) / 100]]);
  ecrireDate_(jou, row, l.date);
  if (cleUid) cache.put(cleUid, String(row), 21600); // mémorisé 6 h
  return { ok: true, row };
}

function modifier_(row, attendu, l) {
  const err = valider_(l);
  if (err) return { ok: false, error: 'INVALIDE', message: err };
  const jou = classeur_().getSheetByName('Journal');
  const r = jou.getRange(row, 1, 1, 4);
  if (!identique_(r.getValues()[0], attendu)) return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans le fichier.', detail: { row, fichier: r.getValues()[0].map(String), attendu } };
  r.setValues([[l.mois, l.lib || l.cat, l.cat, Math.round(Number(l.montant) * 100) / 100]]);
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(l.date || ''))) ecrireDate_(jou, row, l.date); // date de la dépense corrigée dans l'appli
  return { ok: true, row };
}

function supprimer_(row, attendu) {
  const jou = classeur_().getSheetByName('Journal');
  const r = jou.getRange(row, 1, 1, 4);
  if (!identique_(r.getValues()[0], attendu)) return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans le fichier.', detail: { row, fichier: r.getValues()[0].map(String), attendu } };
  if (row > 5) jou.deleteRow(row); else { r.clearContent(); if (jou.getMaxColumns() >= COL_DATE) jou.getRange(row, COL_DATE).clearContent(); }
  return { ok: true, row, decale: row > 5 };
}

/* ---------- Échéancier ---------- */

const FREQS_ = ['Mensuel', 'Trimestriel', 'Semestriel', 'Annuel'];

function echIdentique_(v, a) {
  return a && String(v[0]).trim() === String(a.lib) && String(v[1]).trim() === String(a.cat) &&
    Math.abs(Number(v[2]) - Number(a.montant)) < 0.005;
}

function echEnregistrer_(row, attendu, e) {
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
  if (!row) {
    const a = ech.getRange(5, 1, finEch_(ech) - 4, 1).getValues();
    let i = a.findIndex(r => String(r[0]).trim() === '');
    if (i < 0) return { ok: false, error: 'PLEIN', message: "L'Échéancier est plein (" + a.length + " lignes). Supprime une ligne terminée." };
    row = 5 + i;
  } else {
    if (row < 5 || row > finEch_(ech)) return { ok: false, error: 'INVALIDE', message: 'Ligne hors Échéancier.' };
    if (!echIdentique_(ech.getRange(row, 1, 1, 3).getValues()[0], attendu)) return { ok: false, error: 'CHANGE', message: "La ligne a changé dans l'Échéancier." };
  }
  ech.getRange(row, 1, 1, 10).setValues([ligne]);
  return { ok: true, row };
}

function echSupprimer_(row, attendu) {
  const ech = classeur_().getSheetByName('Échéancier');
  if (row < 5 || row > finEch_(ech)) return { ok: false, error: 'INVALIDE', message: 'Ligne hors Échéancier.' };
  if (!echIdentique_(ech.getRange(row, 1, 1, 3).getValues()[0], attendu)) return { ok: false, error: 'CHANGE', message: "La ligne a changé dans l'Échéancier." };
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

function avEnregistrer_(row, attendu, a) {
  const ss = classeur_(), sh = ss.getSheetByName('À venir');
  if (!sh) return { ok: false, error: 'INVALIDE', message: 'Feuille « À venir » introuvable.' };
  const err = valider_({ mois: a.mois, lib: a.lib, cat: a.cat, montant: a.montant }, true);
  if (err) return { ok: false, error: 'INVALIDE', message: err };
  if (!row) {
    const v = sh.getRange(4, 1, finAv_(sh) - 3, 4).getValues();
    const i = v.findIndex(r => String(r[1]).trim() === '' && r[3] === '');
    if (i < 0) return { ok: false, error: 'PLEIN', message: 'La feuille « À venir » est pleine.' };
    row = 4 + i;
  } else if (!avIdentique_(sh.getRange(row, 1, 1, 4).getValues()[0], attendu)) {
    return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans « À venir ».' };
  }
  // Écrit cellule par cellule : en cas de souci (validation, format…), le message dit laquelle.
  const vals = [a.mois, String(a.lib).trim(), a.cat, Math.round(Number(a.montant) * 100) / 100];
  // Le mois est forcé en texte : sinon Sheets transforme « Novembre 2026 » en date et le menu le refuse.
  sh.getRange(row, 1).setNumberFormat('@');
  ['Mois', 'Libellé', 'Catégorie', 'Montant'].forEach((nom, k) => {
    try { sh.getRange(row, k + 1).setValue(vals[k]); } catch (err) { throw new Error('« À venir », colonne ' + nom + ' : ' + err.message); }
  });
  sh.getRange(row, 5).setValue(false);
  sh.getRange(row, 6).setValue(a.jour ? Math.round(Number(a.jour)) : '');
  return { ok: true, row };
}

function avSupprimer_(row, attendu) {
  const sh = classeur_().getSheetByName('À venir');
  if (!avIdentique_(sh.getRange(row, 1, 1, 4).getValues()[0], attendu)) return { ok: false, error: 'CHANGE', message: 'La ligne a changé dans « À venir ».' };
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