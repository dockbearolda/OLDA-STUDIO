// ============================================================
//  OLDA — Google Apps Script  |  Enregistrement des commandes
//  À déployer comme Application Web (voir instructions bas de page)
// ============================================================

const SHEET_NAME = 'Commandes';
const TIMEZONE   = 'Europe/Paris';

// ─── Ordre des colonnes (NE PAS modifier sans réinitialiser les en-têtes) ───
const HEADERS = [
  'N° Commande',        // col  1
  'Date',               // col  2
  'Client',             // col  3
  'Téléphone',          // col  4  → forcé en texte pour conserver le 0
  'Collection',         // col  5
  'Référence',          // col  6
  'Taille',             // col  7
  'Couleur T-shirt',    // col  8
  'Logo Avant',         // col  9
  'Couleur Logo Av.',   // col 10
  'Logo Arrière',       // col 11
  'Couleur Logo Ar.',   // col 12
  'Note',               // col 13
  'Prix T-shirt (€)',   // col 14
  'Personnalisation (€)',// col 15
  'Total (€)',          // col 16
  'Statut paiement',    // col 17
  'Acompte (€)',        // col 18
  'Horodatage'          // col 19
];

// ─── Couleurs en-têtes ───────────────────────────────────────
const COLOR_HEADER_BG = '#1C1C2E';
const COLOR_HEADER_FG = '#FFFFFF';

// ─── Couleurs par statut de paiement (cellule Statut) ────────
const COLOR_PAYE_OUI   = '#C6EFCE';  // vert
const COLOR_PAYE_ACPTE = '#FFEB9C';  // jaune
const COLOR_PAYE_NON   = '#FFC7CE';  // rouge

// ─── Couleurs de ligne par collection ────────────────────────
const COLOR_COLLECTION = {
  'HOMME'      : '#DBEAFE',  // bleu clair
  'FEMME'      : '#FCE7F3',  // rose clair
  'ENFANT'     : '#D1FAE5',  // vert clair
  'ACCESSOIRE' : '#EDE9FE',  // violet clair
  'DEFAULT'    : '#F8F9FA'   // gris neutre
};

// ─────────────────────────────────────────────────────────────
//  POINT D'ENTRÉE  —  reçoit les commandes du formulaire
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var raw = e.postData ? e.postData.contents : '';
    if (!raw) throw new Error('Corps de requête vide');

    var data  = JSON.parse(raw);
    var sheet = getOrCreateSheet_();

    var ts = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

    // ── Ordre strict calé sur HEADERS ────────────────────────
    var row = [
      data.commande            || '',  //  1  N° Commande
      data.date                || '',  //  2  Date
      data.nom                 || '',  //  3  Client
      data.telephone           || '',  //  4  Téléphone
      data.collection          || '',  //  5  Collection
      data.reference           || '',  //  6  Référence
      data.taille              || '',  //  7  Taille
      data.couleurTshirt       || '',  //  8  Couleur T-shirt
      data.logoAvant           || '',  //  9  Logo Avant
      data.couleurLogoAvant    || '',  // 10  Couleur Logo Av.
      data.logoArriere         || '',  // 11  Logo Arrière
      data.couleurLogoArriere  || '',  // 12  Couleur Logo Ar.
      data.note                || '',  // 13  Note
      data.prixTshirt          || '',  // 14  Prix T-shirt
      data.personnalisation    || '',  // 15  Personnalisation
      data.total               || '',  // 16  Total
      data.paye                || '',  // 17  Statut paiement
      data.acompte             || '',  // 18  Acompte
      ts                               // 19  Horodatage
    ];

    sheet.appendRow(row);

    var lastRow = sheet.getLastRow();

    // Force le téléphone en texte (conserve le 0 initial)
    sheet.getRange(lastRow, 4).setNumberFormat('@');

    formatDataRow_(sheet, lastRow, data.paye, data.collection);

    return jsonResponse_({ status: 'ok', row: lastRow });

  } catch (err) {
    logError_(err);
    return jsonResponse_({ status: 'error', message: err.message });
  }
}

// ─────────────────────────────────────────────────────────────
//  GET  —  vérification que le script est en ligne
// ─────────────────────────────────────────────────────────────
function doGet() {
  return jsonResponse_({ status: 'ok', message: 'OLDA API opérationnelle' });
}

// ─────────────────────────────────────────────────────────────
//  FEUILLE  —  création ou récupération + contrôle des en-têtes
// ─────────────────────────────────────────────────────────────
function getOrCreateSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupHeaders_(sheet);
  } else {
    // Vérifie si les en-têtes correspondent ; les corrige sinon
    var currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    var headersMatch = HEADERS.every(function(h, i) { return h === currentHeaders[i]; });
    if (!headersMatch) {
      setupHeaders_(sheet);
    }
  }

  return sheet;
}

// ─────────────────────────────────────────────────────────────
//  EN-TÊTES  —  écriture + style + formats de colonnes
// ─────────────────────────────────────────────────────────────
function setupHeaders_(sheet) {
  var n = HEADERS.length;

  // En-têtes texte
  sheet.getRange(1, 1, 1, n).setValues([HEADERS]);

  // Style
  var hRange = sheet.getRange(1, 1, 1, n);
  hRange.setBackground(COLOR_HEADER_BG);
  hRange.setFontColor(COLOR_HEADER_FG);
  hRange.setFontWeight('bold');
  hRange.setFontSize(11);
  hRange.setHorizontalAlignment('center');
  hRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // Force la colonne Téléphone (col 4) en texte sur tout le sheet
  sheet.getRange(1, 4, sheet.getMaxRows(), 1).setNumberFormat('@');

  // Largeurs de colonnes
  sheet.setColumnWidth(1,  165);  // N° Commande
  sheet.setColumnWidth(2,  110);  // Date
  sheet.setColumnWidth(3,  160);  // Client
  sheet.setColumnWidth(4,  120);  // Téléphone
  sheet.setColumnWidth(5,  110);  // Collection
  sheet.setColumnWidth(6,  110);  // Référence
  sheet.setColumnWidth(7,   80);  // Taille
  sheet.setColumnWidth(8,  130);  // Couleur T-shirt
  sheet.setColumnWidth(9,  130);  // Logo Avant
  sheet.setColumnWidth(10, 130);  // Couleur Logo Av.
  sheet.setColumnWidth(11, 130);  // Logo Arrière
  sheet.setColumnWidth(12, 130);  // Couleur Logo Ar.
  sheet.setColumnWidth(13, 200);  // Note
  sheet.setColumnWidth(14, 120);  // Prix T-shirt
  sheet.setColumnWidth(15, 145);  // Personnalisation
  sheet.setColumnWidth(16, 100);  // Total
  sheet.setColumnWidth(17, 130);  // Statut paiement
  sheet.setColumnWidth(18, 100);  // Acompte
  sheet.setColumnWidth(19, 155);  // Horodatage
}

// ─────────────────────────────────────────────────────────────
//  MISE EN FORME  —  ligne de données
// ─────────────────────────────────────────────────────────────
function formatDataRow_(sheet, rowIndex, paye, collection) {
  var n        = HEADERS.length;
  var rowRange = sheet.getRange(rowIndex, 1, 1, n);

  // Couleur de fond selon la collection
  var collKey = (collection || '').toUpperCase().trim();
  var bgColor = COLOR_COLLECTION[collKey] || COLOR_COLLECTION['DEFAULT'];
  rowRange.setBackground(bgColor);
  rowRange.setVerticalAlignment('middle');
  sheet.setRowHeight(rowIndex, 28);

  // Couleur de la cellule "Statut paiement" (col 17)
  var payCell   = sheet.getRange(rowIndex, 17);
  var payeUpper = (paye || '').toUpperCase();
  if (payeUpper === 'OUI') {
    payCell.setBackground(COLOR_PAYE_OUI);
    payCell.setFontWeight('bold');
  } else if (payeUpper === 'ACOMPTE') {
    payCell.setBackground(COLOR_PAYE_ACPTE);
    payCell.setFontWeight('bold');
  } else {
    payCell.setBackground(COLOR_PAYE_NON);
    payCell.setFontWeight('bold');
  }
}

// ─────────────────────────────────────────────────────────────
//  RÉINITIALISER LES EN-TÊTES  —  à lancer UNE FOIS manuellement
//  si les colonnes du sheet existant ne correspondent pas
// ─────────────────────────────────────────────────────────────
function reinitialiserEntetes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Feuille "' + SHEET_NAME + '" introuvable.');
    return;
  }
  setupHeaders_(sheet);
  SpreadsheetApp.getUi().alert('En-têtes réinitialisées avec succès.');
}

// ─────────────────────────────────────────────────────────────
//  UTILITAIRES
// ─────────────────────────────────────────────────────────────
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function logError_(err) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Erreurs');
  if (!sheet) sheet = ss.insertSheet('Erreurs');
  var ts = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  sheet.appendRow([ts, err.message, err.stack || '']);
}

// ─────────────────────────────────────────────────────────────
//  FONCTION DE TEST  —  à lancer manuellement depuis l'éditeur
// ─────────────────────────────────────────────────────────────
function testerAvecCommandeFactice() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        commande           : '2026-0217-TestClient',
        date               : '17/02/2026',
        nom                : 'Client Test',
        telephone          : '0600000000',
        collection         : 'Femme',
        reference          : 'F-002',
        taille             : 'M',
        couleurTshirt      : 'Rose',
        logoAvant          : 'OLDA-03',
        couleurLogoAvant   : 'Blanc',
        logoArriere        : '',
        couleurLogoArriere : '',
        note               : 'Test automatique — vérif logo + téléphone + couleur ligne',
        prixTshirt         : '25',
        personnalisation   : '10',
        total              : '35 €',
        paye               : 'ACOMPTE',
        acompte            : '15'
      })
    }
  };

  var result = doPost(fakeEvent);
  Logger.log('Résultat : ' + result.getContent());
}

// ─────────────────────────────────────────────────────────────
//  INSTRUCTIONS DE DÉPLOIEMENT
// ─────────────────────────────────────────────────────────────
//
//  PREMIÈRE INSTALLATION :
//  1. Ouvrir le Google Sheet → Extensions → Apps Script
//  2. Supprimer le code existant, coller TOUT ce fichier
//  3. Sauvegarder (Ctrl+S)
//  4. Exécuter "reinitialiserEntetes" (accepter les autorisations)
//     → Corrige les en-têtes si l'ancien script était différent
//  5. Exécuter "testerAvecCommandeFactice" → vérifier que la ligne
//     apparaît avec la bonne couleur et le 0 du téléphone conservé
//  6. Déployer → Nouveau déploiement
//       Type         : Application Web
//       Exécuter en  : Moi (votre compte Google)
//       Accès        : Tout le monde
//  7. Copier l'URL générée
//  8. Dans index.html ligne 705, mettre à jour la constante API :
//       const API = "COLLER_L_URL_ICI";
//
//  COULEURS DES LIGNES :
//  - Homme      → bleu clair
//  - Femme      → rose clair
//  - Enfant     → vert clair
//  - Accessoire → violet clair
//  - Statut paiement : vert = OUI | jaune = ACOMPTE | rouge = NON
//
// ─────────────────────────────────────────────────────────────
