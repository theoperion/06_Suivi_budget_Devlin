// ============================================================================================
// SAUVEGARDE DANS LE GOOGLE DRIVE DE LA PERSONNE (phase 2, mode « données sur ce téléphone »)
// - Connexion Google dans l'appli (Google Identity Services), accès limité aux fichiers CRÉÉS PAR L'APPLI (drive.file) :
//   l'appli ne voit rien d'autre du Drive. Aucun serveur entre le téléphone et le Drive : personne d'autre ne voit les données.
// - Deux fichiers dans le Drive : un Google Sheet « Jauge · sauvegarde » (mêmes onglets que le fichier budget, valeurs
//   seulement) pour lire ses données, et un fichier technique (JSON) pour tout restaurer sur un nouveau téléphone.
// - Google ne donne l'accès que pour 1 h : pendant ce temps la sauvegarde suit chaque changement ; ensuite il faut un geste
//   (bouton « Sauvegarder ») pour le renouveler.
// Chargé seulement en mode téléphone (comme moteur.js).
// ============================================================================================
(function () {
  'use strict';
  const SCOPE = 'https://www.googleapis.com/auth/drive.file';
  const NOM_SHEET = 'Jauge · sauvegarde', NOM_JSON = 'Jauge · sauvegarde technique (ne pas modifier).json';
  const ONGLETS = ['Journal', 'Échéancier', 'Paramètres', 'À venir'];
  const pad = n => String(n).padStart(2, '0');

  let gisPromesse = null;
  function chargerGIS() {
    if (window.google && google.accounts && google.accounts.oauth2) return Promise.resolve();
    return gisPromesse || (gisPromesse = new Promise((ok, ko) => {
      const s = document.createElement('script'); s.src = 'https://accounts.google.com/gsi/client'; s.async = true;
      s.onload = () => ok(); s.onerror = () => { gisPromesse = null; ko(new Error('Connexion à Google impossible (pas de réseau ?)')); };
      document.head.appendChild(s);
    }));
  }

  // Grille de valeurs d'un onglet imité (moteur.js) : mêmes lignes et colonnes que le fichier, dates en jj/mm/aaaa.
  function grille(o, nom) {
    const der = Object.keys(o.v).reduce((m, k) => Math.max(m, +k), 0), larg = o.c;
    const val = x => x && typeof x === 'object' && x.$d ? x.$d.slice(8, 10) + '/' + x.$d.slice(5, 7) + '/' + x.$d.slice(0, 4) : (x === undefined || x === null ? '' : x);
    const type = g => !g ? '' : g === 'Revenus' ? 'Revenu' : g === 'Épargne' ? 'Épargne' : 'Dépense';
    const G = [];
    for (let r = 1; r <= der; r++) {
      const L = (o.v[r] || []).map(val); while (L.length < larg) L.push('');
      if (nom === 'Paramètres' && r >= 10 && r <= 40) L[2] = type(String(L[1] || '').trim()); // colonne « Type » (formule dans le fichier)
      G.push(L.slice(0, larg));
    }
    return G.length ? G : [['']];
  }

  function Sauvegarde(opt) {
    // opt : { clientId, LS, moteur: () => Promise<MoteurTelephone> }
    const B = this, LS = opt.LS;
    let jeton = null, expire = 0, client = null, attente = null;
    B.etat = () => LS.get('svg', null);              // { sheet, sheetUrl, json, ids: {onglet: id}, lignes: {onglet: n}, le, rev, compte }
    B.connecte = () => !!(jeton && Date.now() < expire - 60000);
    B.aJour = rev => { const e = B.etat(); return !!(e && e.rev === rev); };

    // Demander l'accès (avec un geste de la personne : bouton). Première fois : écran de choix du compte et d'accord.
    B.connecter = async (premiere) => {
      await chargerGIS();
      return new Promise((ok, ko) => {
        const fin = r => { if (r && r.access_token) { jeton = r.access_token; expire = Date.now() + (Number(r.expires_in) || 3600) * 1000; ok(true); } else ko(new Error(r && (r.error_description || r.error) || 'accès refusé')); };
        client = google.accounts.oauth2.initTokenClient({ client_id: opt.clientId, scope: SCOPE, callback: fin, error_callback: e => ko(new Error(e && (e.message || e.type) || 'fenêtre Google fermée')) });
        const e = B.etat();
        client.requestAccessToken(Object.assign({ prompt: premiere ? 'consent' : '' }, e && e.compte ? { login_hint: e.compte } : {}));
      });
    };

    async function g(url, init) {
      const r = await fetch(url, Object.assign({}, init, { headers: Object.assign({ Authorization: 'Bearer ' + jeton }, (init && init.headers) || {}) }));
      if (r.status === 401) { jeton = null; throw new Error('accès Google expiré'); }
      if (!r.ok) { let t = ''; try { t = (await r.json()).error.message; } catch (e) {} throw new Error('Google ' + r.status + (t ? ' · ' + t : '')); }
      return r.status === 204 ? null : r.json();
    }

    // Une sauvegarde complète : fichier technique puis Google Sheet. Rien n'est écrit dans le téléphone sauf la date.
    B.sauver = async () => {
      if (!B.connecte()) throw new Error('pas connecté au Drive');
      const M = await opt.moteur(), C = M.instantane(); if (!C) throw new Error('aucune donnée sur ce téléphone');
      const e = Object.assign({ ids: {}, lignes: {} }, B.etat() || {});
      // 1) fichier technique (JSON complet)
      const corps = JSON.stringify({ app: 'jauge', format: 1, sauve: new Date().toISOString(), classeur: C });
      const json = await (e.json
        ? g('https://www.googleapis.com/upload/drive/v3/files/' + e.json + '?uploadType=media&fields=id', { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: corps })
          .catch(err => { if (/ 404/.test(err.message)) { e.json = null; return null; } throw err; })
        : null) || (e.json ? { id: e.json } : await creerJSON(corps));
      e.json = json.id;
      // 2) Google Sheet lisible (créé une fois, puis vidé et réécrit)
      if (e.sheet) { try { await g('https://sheets.googleapis.com/v4/spreadsheets/' + e.sheet + '?fields=spreadsheetId'); } catch (err) { if (/ 404/.test(err.message)) e.sheet = null; else throw err; } }
      if (!e.sheet) {
        const s = await g('https://sheets.googleapis.com/v4/spreadsheets', { method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ properties: { title: NOM_SHEET, locale: 'fr_FR', timeZone: 'Europe/Paris' }, sheets: ONGLETS.map(t => ({ properties: { title: t, gridProperties: { rowCount: 1000, columnCount: 26 } } })) }) });
        e.sheet = s.spreadsheetId; e.sheetUrl = s.spreadsheetUrl; e.ids = {}; e.lignes = {};
        s.sheets.forEach(x => { e.ids[x.properties.title] = x.properties.sheetId; e.lignes[x.properties.title] = x.properties.gridProperties.rowCount; });
      }
      const G = {}; ONGLETS.forEach(t => { G[t] = C.onglets[t] ? grille(C.onglets[t], t) : [['']]; });
      const agrandir = ONGLETS.filter(t => G[t].length > (e.lignes[t] || 1000) && e.ids[t] != null).map(t => {
        e.lignes[t] = G[t].length + 500;
        return { updateSheetProperties: { properties: { sheetId: e.ids[t], gridProperties: { rowCount: e.lignes[t], columnCount: 26 } }, fields: 'gridProperties.rowCount,gridProperties.columnCount' } };
      });
      if (agrandir.length) await g('https://sheets.googleapis.com/v4/spreadsheets/' + e.sheet + ':batchUpdate', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ requests: agrandir }) });
      const q = n => "'" + n.replace(/'/g, "''") + "'";
      await g('https://sheets.googleapis.com/v4/spreadsheets/' + e.sheet + '/values:batchClear', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ranges: ONGLETS.map(q) }) });
      await g('https://sheets.googleapis.com/v4/spreadsheets/' + e.sheet + '/values:batchUpdate', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ valueInputOption: 'RAW', data: ONGLETS.map(t => ({ range: q(t) + '!A1', values: G[t] })) }) });
      e.le = new Date().toISOString(); e.rev = C.rev;
      LS.set('svg', e);
      return e;
    };
    async function creerJSON(corps) {
      const sep = 'jauge' + Math.random().toString(36).slice(2);
      const meta = { name: NOM_JSON, mimeType: 'application/json', appProperties: { jauge: 'sauvegarde' } };
      const body = '--' + sep + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) + '\r\n--' + sep + '\r\nContent-Type: application/json\r\n\r\n' + corps + '\r\n--' + sep + '--';
      return g('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', { method: 'POST', headers: { 'Content-Type': 'multipart/related; boundary=' + sep }, body });
    }

    // Retrouver la dernière sauvegarde technique dans le Drive (nouveau téléphone) : { sauve, classeur } ou null
    B.lireDrive = async () => {
      if (!B.connecte()) throw new Error('pas connecté au Drive');
      const q = encodeURIComponent("appProperties has { key='jauge' and value='sauvegarde' } and trashed=false");
      const L = await g('https://www.googleapis.com/drive/v3/files?q=' + q + '&orderBy=modifiedTime desc&pageSize=1&fields=files(id,modifiedTime)');
      if (!L.files || !L.files.length) return null;
      const d = await g('https://www.googleapis.com/drive/v3/files/' + L.files[0].id + '?alt=media');
      if (!d || d.app !== 'jauge' || !d.classeur || !d.classeur.onglets) throw new Error('sauvegarde illisible');
      return Object.assign(d, { id: L.files[0].id });
    };

    // Sauvegarde automatique après un changement : regroupée (20 s), seulement si l'accès Google est encore valable.
    B.planifier = (fait) => {
      clearTimeout(attente);
      attente = setTimeout(async () => {
        if (!B.connecte()) { if (fait) fait(false, 'accès Google à renouveler'); return; }
        try { await B.sauver(); if (fait) fait(true); } catch (err) { if (fait) fait(false, err.message); }
      }, opt.delai == null ? 20000 : opt.delai);
    };
  }
  window.SauvegardeDrive = Sauvegarde;
})();
