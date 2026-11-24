// ============================================================================================
// MOTEUR TÉLÉPHONE (phase 2, 24/11/2026) : les données restent sur le téléphone, sans script Google.
// Principe : le MÊME script que le fichier Google (apps-script/Code.js) tourne ici, dans le téléphone,
// sur une imitation du tableur (onglets Paramètres, Journal, Échéancier, À venir) gardée dans IndexedDB.
// Les règles d'écriture (lignes, anti-doublon, contrôles « la ligne a changé »…) sont donc exactement les mêmes.
// Une écriture n'est confirmée à l'appli qu'une fois enregistrée dans le téléphone ; sinon tout est remis
// comme avant et l'appli garde l'envoi en attente (réponse OCCUPE : nouvel essai automatique).
// ============================================================================================
(function () {
  'use strict';
  const MOIS = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
  const pad = n => String(n).padStart(2, '0');

  // ---------- Valeurs : les dates sont gardées en { $d: 'aaaa-mm-jjThh:mm:ss' } (heure locale) ----------
  const enJSON = v => v instanceof Date ? { $d: v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate()) + 'T' + pad(v.getHours()) + ':' + pad(v.getMinutes()) + ':' + pad(v.getSeconds()) } : v;
  const deJSON = v => v && typeof v === 'object' && typeof v.$d === 'string' ? new Date(+v.$d.slice(0, 4), +v.$d.slice(5, 7) - 1, +v.$d.slice(8, 10), +v.$d.slice(11, 13) || 0, +v.$d.slice(14, 16) || 0, +v.$d.slice(17, 19) || 0) : v;
  const affiche = v => v instanceof Date ? pad(v.getDate()) + '/' + pad(v.getMonth() + 1) + '/' + v.getFullYear() : v === true ? 'TRUE' : v === false ? 'FALSE' : String(v);

  // Type d'une catégorie (colonne C des Paramètres) : formule dans le fichier. Ici, le type importé du fichier est gardé tel quel ;
  // il n'est déduit (groupe + nom) que pour une catégorie créée, renommée ou changée de groupe sur le téléphone.
  // Même règle que la formule du fichier (Paramètres!C10) : =IF(A="";"";IF(B="Revenus";"Revenu";IF(B="Épargne";IF(REGEXMATCH(LOWER(A);"^retrait");"Retrait";"Épargne");"Dépense")))
  const typeDe = (g, nom) => !nom ? '' : g === 'Revenus' ? 'Revenu' : g === 'Épargne' ? (/^retrait/.test(String(nom).toLowerCase()) ? 'Retrait' : 'Épargne') : 'Dépense';

  // ---------- Une feuille imitée ----------
  function Feuille(nom, d) {
    // d = { n: lignes, c: colonnes, v: { "ligne": [valeurs…] } (lignes non vides seulement), f: formats, F: formules }
    const sh = this; sh.nom = nom; sh.d = d;
    const cell = (r, c) => { const L = d.v[r]; const x = L ? L[c - 1] : undefined; return x === undefined || x === null ? '' : deJSON(x); };
    const ecrire = (r, c, x) => {
      if (r < 1 || c < 1 || r > d.n || c > d.c) throw new Error('Plage hors de la feuille « ' + nom + ' » (' + r + ', ' + c + ')');
      const L = d.v[r] || (d.v[r] = []);
      while (L.length < c) L.push('');
      L[c - 1] = x === undefined || x === null ? '' : enJSON(x);
      // Paramètres : le nom ou le groupe d'une catégorie change → son type (colonne C) est recalculé
      if (nom === 'Paramètres' && r >= 10 && r <= 40 && (c === 1 || c === 2)) { while (L.length < 3) L.push(''); L[2] = typeDe(String(L[1] || '').trim(), String(L[0] || '').trim()); }
      if (L.every(y => y === '')) delete d.v[r];
    };
    const lire = (r, c) => cell(r, c);
    sh.getName = () => nom;
    sh.getMaxRows = () => d.n;
    sh.getMaxColumns = () => d.c;
    sh.getLastRow = () => { let m = 0; Object.keys(d.v).forEach(k => { if (+k > m) m = +k; }); return m; };
    sh.hideSheet = () => sh;
    sh.appendRow = vals => { const r = sh.getLastRow() + 1; if (r > d.n) d.n = r; if (vals.length > d.c) d.c = vals.length; vals.forEach((x, j) => ecrire(r, j + 1, x)); return sh; };
    sh.insertRowsAfter = (a, n) => {
      const v = {}; Object.keys(d.v).forEach(k => { v[+k > a ? +k + n : k] = d.v[k]; }); d.v = v;
      decalerCles(d, 'f', a, n); decalerCles(d, 'F', a, n); d.n += n; return sh;
    };
    sh.insertColumnsAfter = (a, n) => { Object.keys(d.v).forEach(k => { const L = d.v[k]; if (L.length > a) L.splice(a, 0, ...Array(n).fill('')); }); d.c += n; return sh; };
    sh.deleteRow = r => sh.deleteRows(r, 1);
    sh.deleteRows = (r, n) => {
      const v = {}; Object.keys(d.v).forEach(k => { k = +k; if (k < r) v[k] = d.v[k]; else if (k >= r + n) v[k - n] = d.v[k]; }); d.v = v;
      supprimerCles(d, 'f', r, n); supprimerCles(d, 'F', r, n); d.n = Math.max(1, d.n - n); return sh;
    };
    sh.getRange = (r, c, nr, nc) => {
      if (typeof r === 'string') {
        const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+)?)?$/.exec(r); if (!m) throw new Error('Plage inconnue : ' + r);
        const C = s => s.split('').reduce((a, x) => a * 26 + x.charCodeAt(0) - 64, 0);
        c = C(m[1]); const r0 = +m[2]; nc = m[3] ? C(m[3]) - c + 1 : 1; nr = m[3] ? (m[4] ? +m[4] : d.n) - r0 + 1 : 1; r = r0;
      }
      nr = nr || 1; nc = nc || 1;
      const R = {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => lire(r + i, c + j))),
        getDisplayValues: () => R.getValues().map(L => L.map(affiche)),
        getValue: () => lire(r, c),
        getDisplayValue: () => affiche(lire(r, c)),
        setValues: vals => {
          if (vals.length !== nr || vals.some(L => L.length !== nc)) throw new Error('Nombre de lignes ou colonnes incorrect');
          vals.forEach((L, i) => L.forEach((x, j) => ecrire(r + i, c + j, x))); return R;
        },
        setValue: x => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) ecrire(r + i, c + j, x); return R; },
        clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { if (r + i <= d.n && c + j <= d.c) ecrire(r + i, c + j, ''); } return R; },
        setNumberFormat: f => { d.f = d.f || {}; for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) d.f[(r + i) + ':' + (c + j)] = f; return R; },
        getFormula: () => (d.F && d.F[r + ':' + c]) || '',
        setFormula: f => { d.F = d.F || {}; d.F[r + ':' + c] = f; return R; }
      };
      return R;
    };
  }
  function decalerCles(d, cle, a, n) { if (!d[cle]) return; const o = {}; Object.keys(d[cle]).forEach(k => { const [r, c] = k.split(':').map(Number); o[(r > a ? r + n : r) + ':' + c] = d[cle][k]; }); d[cle] = o; }
  function supprimerCles(d, cle, r0, n) { if (!d[cle]) return; const o = {}; Object.keys(d[cle]).forEach(k => { const [r, c] = k.split(':').map(Number); if (r < r0) o[k] = d[cle][k]; else if (r >= r0 + n) o[(r - n) + ':' + c] = d[cle][k]; }); d[cle] = o; }

  // ---------- Le classeur imité : { onglets: { nom: données }, props: {}, debut: 'aaaa-mm-jj'|null, rev: n } ----------
  function services(C) {
    const feuilles = {};
    Object.keys(C.onglets).forEach(n => { feuilles[n] = new Feuille(n, C.onglets[n]); });
    const classeur = {
      getSheetByName: n => feuilles[n] || null,
      insertSheet: n => { C.onglets[n] = { n: 5, c: 6, v: {} }; return (feuilles[n] = new Feuille(n, C.onglets[n])); },
      getSpreadsheetTimeZone: () => 'Europe/Paris',
      getRangeByName: n => n === 'DebutEcheancier' ? { getValue: () => C.debut ? new Date(+C.debut.slice(0, 4), +C.debut.slice(5, 7) - 1, +C.debut.slice(8, 10)) : '' } : null,
      getUrl: () => '',
      getName: () => 'Ce téléphone'
    };
    const props = {
      getProperty: k => (Object.prototype.hasOwnProperty.call(C.props, k) ? C.props[k] : null),
      setProperty: (k, v) => { C.props[k] = String(v); return props; },
      deleteProperty: k => { delete C.props[k]; return props; }
    };
    return {
      SpreadsheetApp: { getActiveSpreadsheet: () => classeur, openById: () => classeur },
      PropertiesService: { getScriptProperties: () => props },
      CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
      LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) },
      Utilities: { formatDate: (d, tz, f) => f.replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate())) },
      ContentService: { createTextOutput: t => ({ setMimeType: () => t }), MimeType: { JSON: 'json' } },
      Logger: { log: () => {} },
      console: { log: () => {}, error: () => {}, warn: () => {} }
    };
  }

  // Le script Google, chargé une fois ; ses variables « globales » sont remises à zéro à chaque demande,
  // comme dans Apps Script où chaque demande repart de zéro.
  function compiler(code) {
    const noms = ['SpreadsheetApp', 'PropertiesService', 'CacheService', 'LockService', 'Utilities', 'ContentService', 'Logger', 'console'];
    // Chaque service renvoie à celui du classeur en cours (choisi juste avant chaque demande)
    const f = new Function('$S', noms.map(n => 'const ' + n + ' = new Proxy({}, { get: (_, k) => $S.get().' + n + '[k] });').join('\n') + '\n' + code +
      '\n;return { traiter: function (req) { SS_ = null; UIDS_ = null; UIDS_SALE_ = false; return traiter_(req); }, version: VERSION_API };');
    return f;
  }

  // ---------- Alertes « anomalies » (dans le fichier : formule Journal!A3 et colonne L « Contrôle »), mêmes textes.
  // Seules celles qui peuvent arriver avec des saisies faites par l'appli : doublon possible, déjà dans l'Échéancier,
  // montant inhabituel, incohérences de l'Échéancier, solde au 1er janvier, rappels de début de mois / d'année.
  // L'année vérifiée est l'année en cours (dans le fichier : l'année choisie dans Paramètres).
  const ACTIONS = { 'Doublon possible': "efface l'une des deux si erreur, sinon change un peu le libellé", "Déjà dans l'Échéancier ?": "si c'est le même paiement, efface la ligne du Journal",
    'Montant inhabituel': "vérifie la virgule ; si le montant est juste, ajoute ✓ à la fin du libellé", 'Dates incohérentes': "corrige « Depuis » / « Jusqu'en »",
    'Mois hors de 1 à 12': 'mets un numéro de mois entre 1 et 12', 'Paiement sans mois de départ': 'remplis « 1er mois » (ex. 8 pour août)',
    "Doublon dans l'Échéancier": "efface l'une des deux lignes si c'est le même paiement", 'Solde au 1er janvier': 'saisis le solde de ton relevé au 1er janvier' };
  function montantEch(L, y, m, debut) { // colonnes M:X de l'Échéancier : montant d'une ligne pour un mois donné
    const [lib, , mt, freq, , f, g, h, i, pause] = L;
    if (!lib || pause === true || typeof mt !== 'number') return 0;
    if (debut && new Date(y, m - 1, 1) < debut) return 0;
    if (h !== '' && h != null && y < h) return 0; if (i !== '' && i != null && y > i) return 0;
    const f1 = Math.max(1, +f || 0), pas = { Mensuel: 1, Trimestriel: 3, Semestriel: 6 }[freq] || 12;
    const c1 = (h !== '' && h != null && y > h) || m >= f1, c2 = (i !== '' && i != null && y < i) || m <= (g === '' || g == null ? 12 : g);
    return c1 && c2 && ((((m - f1) % pas) + pas) % pas) === 0 ? mt : 0;
  }
  function anomalies(C, auj) {
    auj = auj || new Date();
    const val = x => (x && typeof x === 'object' && x.$d) ? x : x;
    const jou = C.onglets['Journal'], ech = C.onglets['Échéancier'], par = C.onglets['Paramètres'];
    const groupe = {}; for (let r = 10; r <= 40; r++) { const L = par.v[r]; if (L && L[0]) groupe[String(L[0]).trim()] = String(L[1] || '').trim(); }
    const Y = auj.getFullYear(), M = auj.getMonth() + 1, mcur = MOIS[M - 1] + ' ' + Y, mprev = M > 1 ? MOIS[M - 2] + ' ' + Y : 'Décembre ' + (Y - 1);
    const debut = C.debut ? new Date(+C.debut.slice(0, 4), +C.debut.slice(5, 7) - 1, +C.debut.slice(8, 10)) : null;
    const lignes = Object.keys(jou.v).map(Number).filter(r => r >= 5).sort((a, b) => a - b).map(r => ({ r, L: jou.v[r] })).filter(x => x.L[0] !== '' && x.L[0] != null);
    const E = Object.keys(ech.v).map(Number).filter(r => r >= 5).sort((a, b) => a - b).map(r => ({ r, L: (ech.v[r] || []).concat(Array(10).fill('')).slice(0, 10) })).filter(x => x.L[0] !== '' || x.L[2] !== '');
    const res = [];
    // Doublon possible (mois en cours et précédent)
    const sel = lignes.filter(x => x.L[0] === mcur || x.L[0] === mprev), cle = x => x.L[0] + '|' + x.L[1] + '|' + x.L[2] + '|' + x.L[3], nb = {};
    sel.forEach(x => { nb[cle(x)] = (nb[cle(x)] || 0) + 1; });
    sel.filter(x => nb[cle(x)] > 1).forEach(x => res.push(['Doublon possible', 'Journal, ligne ' + x.r]));
    // Déjà dans l'Échéancier ? (colonne L : même catégorie, même mois de l'année, montant à 10 % près ; libellé finissant par ✓ ignoré)
    lignes.forEach(x => {
      const p = String(x.L[0]).split(' '), mo = MOIS.indexOf(p[0]) + 1, an = +p[1];
      if (an !== Y || mo < 1 || /✓\s*$/.test(String(x.L[1] || ''))) return;
      const d = typeof x.L[3] === 'number' ? x.L[3] : 0;
      if (E.some(e => e.L[1] === x.L[2] && (() => { const a = montantEch(e.L, Y, mo, debut); return a > 0 && Math.abs(d - a) <= 0.1 * a; })())) res.push(["Déjà dans l'Échéancier ?", 'Journal, ligne ' + x.r]);
    });
    // Échéancier
    E.forEach(e => { const [a, , , d, , f, g, h, i] = e.L; if (a === '') return; const R = 'Échéancier, ligne ' + e.r;
      if (h !== '' && i !== '' && +i < +h) res.push(['Dates incohérentes', R]);
      if ((f !== '' && (+f < 1 || +f > 12)) || (g !== '' && (+g < 1 || +g > 12))) res.push(['Mois hors de 1 à 12', R]);
      if (d !== 'Mensuel' && d !== '' && f === '') res.push(['Paiement sans mois de départ', R]); });
    // Montant inhabituel (≥ 1 000 €, mois en cours et précédent, hors Revenus et Épargne, libellé sans ✓)
    lignes.forEach(x => { const d = x.L[3], g = groupe[String(x.L[2]).trim()];
      if ((x.L[0] === mcur || x.L[0] === mprev) && typeof d === 'number' && d >= 1000 && g !== 'Revenus' && g !== 'Épargne' && !/✓\s*$/.test(String(x.L[1] || ''))) res.push(['Montant inhabituel : ' + Math.round(d) + ' €', 'Journal, ligne ' + x.r]); });
    const cE = {}; E.forEach(e => { if (e.L[0] !== '') { const k = e.L[0] + '|' + e.L[1] + '|' + e.L[2]; cE[k] = (cE[k] || 0) + 1; } });
    E.forEach(e => { if (e.L[0] !== '' && cE[e.L[0] + '|' + e.L[1] + '|' + e.L[2]] > 1) res.push(["Doublon dans l'Échéancier", 'Échéancier, ligne ' + e.r]); });
    // Solde au 1er janvier de l'année en cours
    let solde = 0; for (let r = 52; r <= 78; r++) { const L = par.v[r]; if (L && String(L[0]) === String(Y)) solde = typeof L[1] === 'number' ? L[1] : 0; }
    if (!solde) res.push(['Solde au 1er janvier ' + Y + ' non saisi', 'Paramètres ③']);
    const k = res.length, act = n => ACTIONS[Object.keys(ACTIONS).find(a => n.startsWith(a))] || 'à vérifier';
    const lst = k === 0 ? '✅ Aucune anomalie dans les saisies' : '⚠ ' + k + ' ANOMALIE(S) À CORRIGER\n' + res.slice(0, 8).map(([n, o]) => '•  ' + n + '  —  ' + o + '  →  ' + act(n)).join('\n') + (k > 8 ? '\n•  … et ' + (k - 8) + ' autre(s) : corrige d\'abord celles-ci' : '');
    const rap = [lst];
    if (auj.getDate() <= 4) rap.push('📅 Rappel (pas une anomalie) : nouveau mois, salaire arrivé, contrôle ton solde bancaire dans Paramètres ②');
    if (M === 1 && auj.getDate() <= 4) rap.push('📅 Rappel nouvelle année : revois tes budgets (Paramètres ①)');
    void val; return rap.join('\n');
  }

  // ---------- Mois et années à venir : la liste des mois (Paramètres, colonne N) et les années des soldes (A52:A90)
  // vont toujours jusqu'à décembre de l'an prochain (dans le fichier, c'était fait à la main ou par formule).
  function etendreMois(C) {
    const par = C.onglets['Paramètres']; if (!par) return false;
    const fin = new Date().getFullYear() + 1; let change = false;
    const liste = []; for (let r = 4; r <= par.n; r++) { const L = par.v[r]; const x = L && L[13]; if (x) liste.push([r, String(x)]); }
    const der = liste.length ? liste[liste.length - 1] : null;
    let y, m, r;
    if (der) { const k = der[1].split(' '); m = MOIS.indexOf(k[0]); y = +k[1]; r = der[0]; } else { y = new Date().getFullYear(); m = -1; r = 3; }
    if (m < 0 && der) return false; // liste illisible : on ne touche à rien
    while (y < fin || (y === fin && m < 11)) {
      m++; if (m > 11) { m = 0; y++; } r++;
      if (r > par.n) par.n = r;
      const L = par.v[r] || (par.v[r] = []); while (L.length < 14) L.push(''); L[13] = MOIS[m] + ' ' + y; change = true;
    }
    const ans = []; for (let k = 52; k <= 90; k++) { const L = par.v[k]; if (L && L[0] !== '' && L[0] != null) ans.push(+L[0]); }
    let libre = 52; while (libre <= 90 && par.v[libre] && par.v[libre][0] !== '' && par.v[libre][0] != null) libre++;
    const maxAn = ans.length ? Math.max(...ans) : new Date().getFullYear() - 1;
    for (let a = maxAn + 1; a <= fin && libre <= 90; a++, libre++) { const L = par.v[libre] || (par.v[libre] = []); while (L.length < 2) L.push(''); L[0] = a; change = true; }
    return change;
  }

  // ---------- Fabriquer le classeur à partir de la réponse « data » du fichier Google (import, une seule fois) ----------
  function depuisDonnees(data, cle) {
    const ligneMax = (L, k, min) => L.reduce((m, x) => Math.max(m, +x[k] || 0), min);
    const par = { n: 90, c: 14, v: {} }, jou = { n: 0, c: 14, v: {} }, ech = { n: 0, c: 10, v: {} }, av = { n: 0, c: 6, v: {} };
    const met = (o, r, c, x) => { const L = o.v[r] || (o.v[r] = []); while (L.length < c) L.push(''); L[c - 1] = enJSON(x); };
    // Paramètres : mois de l'année (I4:I15), liste des mois du Journal (N4…), catégories (A10:E40), contrôle (B43:B45), soldes (A52:B90)
    MOIS.forEach((m, i) => met(par, 4 + i, 9, m));
    (data.mois || []).forEach((m, i) => met(par, 4 + i, 14, m));
    par.n = Math.max(90, 3 + (data.mois || []).length);
    for (let r = 10; r <= 40; r++) { met(par, r, 1, 'Catégorie libre'); met(par, r, 5, 'Annuel'); }
    (data.categories || []).forEach(c => {
      if (!(c.ligne >= 10 && c.ligne <= 40)) throw new Error('Catégorie hors des lignes 10 à 40 : ' + c.nom);
      met(par, c.ligne, 1, c.nom); met(par, c.ligne, 2, c.groupe); met(par, c.ligne, 3, c.type || ''); met(par, c.ligne, 4, typeof c.budget === 'number' ? c.budget : ''); met(par, c.ligne, 5, c.suivi || '');
    });
    const ctl = data.controle || {};
    met(par, 43, 2, ctl.mois || ''); met(par, 44, 2, typeof ctl.solde === 'number' ? ctl.solde : ''); met(par, 45, 2, typeof ctl.especes === 'number' ? ctl.especes : '');
    const ans = new Set(Object.keys(data.soldes || {}));
    (data.mois || []).forEach(m => { const a = String(m).split(' ')[1]; if (a) ans.add(a); });
    [...ans].sort().slice(0, 39).forEach((a, i) => { met(par, 52 + i, 1, +a); const s = (data.soldes || {})[a]; if (typeof s === 'number') met(par, 52 + i, 2, s); });
    // Journal : en-têtes ligne 4, lignes à leur numéro d'origine (les numéros servent aux modifications)
    met(jou, 4, 13, 'Saisi le'); met(jou, 4, 14, 'Payé en');
    (data.journal || []).forEach(l => {
      met(jou, l.row, 1, l.mois); met(jou, l.row, 2, l.lib); met(jou, l.row, 3, l.cat); met(jou, l.row, 4, l.montant);
      if (l.date) met(jou, l.row, 13, new Date(+l.date.slice(0, 4), +l.date.slice(5, 7) - 1, +l.date.slice(8, 10)));
      if (l.especes) met(jou, l.row, 14, 'Espèces'); else if (l.enLigne) met(jou, l.row, 14, 'En ligne');
    });
    jou.n = ligneMax(data.journal || [], 'row', 4) + 100;
    (data.echeancier || []).forEach(e => {
      [e.lib, e.cat, e.montant == null ? '' : e.montant, e.freq, e.jour || '', e.f == null ? '' : e.f, e.g == null ? '' : e.g, e.h == null ? '' : e.h, e.i == null ? '' : e.i, e.pause ? true : '']
        .forEach((x, j) => met(ech, e.ligne, j + 1, x));
    });
    ech.n = Math.max(ligneMax(data.echeancier || [], 'ligne', 4) + 30, 60);
    (data.avenir || []).forEach(a => { [a.mois, a.lib, a.cat, a.montant, a.fait === true, a.jour == null ? '' : a.jour].forEach((x, j) => met(av, a.row, j + 1, x)); });
    av.n = Math.max(ligneMax(data.avenir || [], 'row', 3) + 30, 40);
    const props = { APP_KEY: cle, DEFIS: JSON.stringify(data.defis || {}) };
    Object.keys(data.prefs || {}).forEach(k => { props['PREF_' + k] = JSON.stringify(data.prefs[k]); });
    return { format: 1, onglets: { 'Paramètres': par, 'Journal': jou, 'Échéancier': ech, 'À venir': av }, props, debut: data.debut || null, rev: 1, cree: new Date().toISOString() };
  }

  // ---------- Mémoire durable du téléphone (IndexedDB) ----------
  function base(nom) {
    return new Promise((ok, ko) => {
      const q = indexedDB.open(nom, 1);
      q.onupgradeneeded = () => q.result.createObjectStore('kv');
      q.onsuccess = () => ok(q.result); q.onerror = () => ko(q.error || new Error('IndexedDB indisponible'));
    });
  }
  const lireKV = (db, k) => new Promise((ok, ko) => { const t = db.transaction('kv', 'readonly').objectStore('kv').get(k); t.onsuccess = () => ok(t.result); t.onerror = () => ko(t.error); });
  const ecrireKV = (db, k, v) => ecrireTout(db, { [k]: v });
  // Plusieurs clés dans la même transaction : tout ou rien (le classeur et son numéro de révision ensemble)
  const ecrireTout = (db, o) => new Promise((ok, ko) => { const t = db.transaction('kv', 'readwrite', { durability: 'strict' }); const s = t.objectStore('kv'); Object.keys(o).forEach(k => s.put(o[k], k)); t.oncomplete = () => ok(); t.onerror = () => ko(t.error); t.onabort = () => ko(t.error || new Error('Écriture annulée')); });

  function Moteur(prefixe) {
    const M = this; let db = null, C = null, texte = null, prog = null, file = Promise.resolve(), persiste = false;
    const S$ = { courant: null, get() { return this.courant; } };
    async function ouvrirBase() { return db || (db = await base(prefixe + 'telephone')); }
    async function charger() {
      await ouvrirBase();
      // Toujours la dernière version enregistrée : une autre fenêtre de l'appli (onglet du navigateur + appli installée)
      // a pu écrire entre-temps. On compare le numéro de révision (lecture légère) avant de relire tout le classeur.
      const rev = await lireKV(db, 'rev');
      if (!C || rev === undefined || C.rev !== rev) { const t = await lireKV(db, 'classeur'); C = t ? JSON.parse(t) : null; }
      if (!persiste) { persiste = true; try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) {} } // demander au navigateur de ne jamais effacer
      if (!prog) {
        // Le script : la version servie avec l'appli (mise à jour avec elle), sinon la dernière gardée dans le téléphone
        try { const r = await fetch('apps-script/Code.js', { cache: 'no-cache' }); if (r.ok) { const t = await r.text(); if (/function traiter_\(/.test(t)) { texte = t; ecrireKV(db, 'script', t).catch(() => {}); } } } catch (e) {}
        if (!texte) texte = await lireKV(db, 'script');
        if (!texte) throw new Error('Script introuvable (première ouverture sans réseau ?)');
        prog = compiler(texte)(S$);
      }
    }
    M.existe = async () => { await ouvrirBase(); return !!(await lireKV(db, 'classeur')); };
    M.importer = async (data, cle) => {
      const nouveau = depuisDonnees(data, cle);
      await ouvrirBase(); await ecrireTout(db, { classeur: JSON.stringify(nouveau), rev: nouveau.rev }); C = nouveau;
      try { if (navigator.storage && navigator.storage.persist) await navigator.storage.persist(); } catch (e) {}
      return true;
    };
    // Restaurer une sauvegarde (nouveau téléphone, ou téléphone effacé) : le classeur complet remplace celui du téléphone.
    M.restaurer = async (C2, cle) => {
      if (!C2 || !C2.onglets || !C2.onglets['Journal'] || !C2.onglets['Paramètres'] || !C2.props) throw new Error('sauvegarde illisible');
      const nouveau = JSON.parse(JSON.stringify(C2)); if (cle) nouveau.props.APP_KEY = cle;
      await ouvrirBase(); const r0 = await lireKV(db, 'rev'); nouveau.rev = Math.max(nouveau.rev || 0, r0 || 0) + 1;
      await ecrireTout(db, { classeur: JSON.stringify(nouveau), rev: nouveau.rev }); C = nouveau;
      return true;
    };
    M.revision = () => (C ? C.rev : 0);
    M.instantane = () => (C ? JSON.parse(JSON.stringify(C)) : null);   // pour la sauvegarde
    // Une demande à la fois, dans l'ordre (comme le verrou du script)
    // Entre fenêtres de l'appli : un verrou du navigateur (si disponible) pour qu'une seule écrive à la fois.
    const verrou = fn => (navigator.locks && navigator.locks.request) ? navigator.locks.request(prefixe + 'telephone', fn) : fn();
    M.appel = req => (file = file.then(() => verrou(() => traiterUne(req)), () => verrou(() => traiterUne(req))));
    async function traiterUne(req) {
      await charger();
      if (!C) return { ok: false, error: 'SANS_CLE', message: "Aucune donnée sur ce téléphone." };
      etendreMois(C); // mois ajoutés : enregistrés avec la prochaine écriture (refaits à l'identique d'ici là)
      const avant = JSON.stringify(C);
      let res;
      try { S$.courant = services(C); res = JSON.parse(JSON.stringify(prog.traiter(JSON.parse(JSON.stringify(req))))); }
      catch (err) { C = JSON.parse(avant); return { ok: false, error: 'SERVEUR', message: String(err && err.message || err) }; }
      const apres = JSON.stringify(C);
      if (apres !== avant) {
        C.rev = (C.rev || 0) + 1; C.modifie = new Date().toISOString();
        try { await ecrireTout(db, { classeur: JSON.stringify(C), rev: C.rev }); }
        catch (err) { C = JSON.parse(avant); return { ok: false, error: 'OCCUPE', message: "Pas pu enregistrer sur le téléphone, nouvel essai automatique." }; }
      }
      if (res && res.ok && res.data && req.action === 'data') { try { res.data.anomalies = anomalies(C); } catch (e) { res.data.anomalies = ''; } }
      return res;
    }
  }
  window.MoteurTelephone = Moteur;
  window.MoteurTelephone.depuisDonnees = depuisDonnees;   // pour les tests
  window.MoteurTelephone.anomalies = anomalies;
})();
