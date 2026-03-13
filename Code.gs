/* ============================================================
   GÉNÉRATEUR DE CONVENTIONS D'EXPOSITION — L'HERMINE
   Code.gs — Backend Google Apps Script
   ============================================================ */

// === CONFIGURATION ===
const CONFIG = {
  SPREADSHEET_ID: '1FhRkS0Qb-Q8UjYjvcIxq18hrXcmwgQZXKUDK0xenWak',
  SHEET_NAME: 'Historique',
  DRIVE_FOLDER_NAME: 'Conventions Expositions Hermine',
  HEADERS: [
    'ID', 'Date création', 'Dernière modif.',
    'Artiste - Nom', 'Artiste - Adresse', 'Artiste - SIRET',
    'Artiste2 - Actif', 'Artiste2 - Nom', 'Artiste2 - Adresse', 'Artiste2 - SIRET',
    'Expo - Début', 'Expo - Fin', 'Expo - Lieu',
    'Installation - Date', 'Installation - Heure',
    'Vernissage - Date', 'Vernissage - Heure',
    'Démontage - Date', 'Démontage - Heure',
    'Droits présentation (€)', 'TVA applicable',
    'Transport', 'Transport - Détails',
    'Hébergement', 'Hébergement - Détails',
    'Actions culturelles', 'Actions culturelles - Détails',
    'Date signature', 'Statut', 'Lien PDF', 'Articles personnalisés',
    'Articles standards personnalisés'
  ]
};

// === POINT D'ENTRÉE ===
function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : 'index';
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Conventions Expositions — L\'Hermine')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// === UTILITAIRES SPREADSHEET ===
function getOrCreateSpreadsheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  // Vérifier que la feuille Historique existe, sinon la créer
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1e3a5f')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 180);
  }

  // S'assurer que les en-têtes existent (évolutions)
  try { ensureHeaders_(sheet); } catch(e) {}

  return ss;
}


function ensureHeaders_(sheet) {
  if (!sheet) return;
  var needCols = CONFIG.HEADERS.length;
  var currentLastCol = sheet.getLastColumn();
  var readCols = Math.max(currentLastCol, needCols);
  var existing = sheet.getRange(1, 1, 1, readCols).getValues()[0] || [];
  var hasUpdate = false;

  for (var i = 0; i < needCols; i++) {
    if (!existing[i]) {
      sheet.getRange(1, i + 1).setValue(CONFIG.HEADERS[i]);
      hasUpdate = true;
    }
  }

  if (hasUpdate) {
    sheet.getRange(1, 1, 1, needCols)
      .setFontWeight('bold')
      .setBackground('#1e3a5f')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}


function getSheet_() {
  var ss = getOrCreateSpreadsheet_();
  // Chercher la feuille Historique, sinon prendre la première feuille
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    // Peut-être que la feuille existante a un autre nom (ex: "Feuille 1")
    // Si elle contient des données qui ressemblent à des conventions, on l'utilise
    sheet = ss.getSheets()[0];
  }
  return sheet;
}

function getOrCreateDriveFolder_() {
  const folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
}

// === GÉNÉRATION D'ID ===
function generateId_() {
  return 'CONV-' + Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyyMMdd') + '-' + Math.random().toString(36).substring(2, 6).toUpperCase();
}

// === RÉCUPÉRER TOUTES LES CONVENTIONS ===
function getConventions() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // Lire toutes les colonnes disponibles
  var numCols = sheet.getLastColumn();
  if (numCols < 1) return [];
  var colsToRead = Math.max(numCols, CONFIG.HEADERS.length);
  
  var data;
  try {
    data = sheet.getRange(2, 1, lastRow - 1, colsToRead).getValues();
  } catch(e) {
    data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  }

  var results = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    // Ignorer les lignes vides
    if (!row[0] && !row[3]) continue;
    
    try {
      results.push({
        id: String(row[0] || ''),
        dateCreation: formatDateSafe_(row[1]),
        derniereModif: formatDateSafe_(row[2]),
        artiste: {
          nom: String(row[3] || ''),
          adresse: String(row[4] || ''),
          siret: String(row[5] || '')
        },
        artiste2: {
          actif: String(row[6] || '') === 'Oui',
          nom: String(row[7] || ''),
          adresse: String(row[8] || ''),
          siret: String(row[9] || '')
        },
        expo: {
          debut: formatDateOnly_(row[10]),
          fin: formatDateOnly_(row[11]),
          lieu: String(row[12] || '') || 'Espace Culturel L\u2019Hermine'
        },
        installation: { date: formatDateOnly_(row[13]), heure: formatTimeSafe_(row[14]) },
        vernissage: { date: formatDateOnly_(row[15]), heure: formatTimeSafe_(row[16]) },
        demontage: { date: formatDateOnly_(row[17]), heure: formatTimeSafe_(row[18]) },
        droits: { montant: String(row[19] || ''), tva: String(row[20] || '') },
        transport: { actif: String(row[21] || ''), details: String(row[22] || '') },
        hebergement: { actif: String(row[23] || ''), details: String(row[24] || '') },
        actionsCulturelles: { actif: String(row[25] || ''), details: String(row[26] || '') },
        dateSignature: formatDateOnly_(row[27]),
        statut: String(row[28] || '') || 'Brouillon',
        lienPdf: String(row[29] || ''),
        articlesPerso: row[30] ? String(row[30]) : '[]',
        articlesStd: row[31] ? String(row[31]) : '{}'
      });
    } catch(e) {
      // Ligne mal formatée, on l'ignore pas — on met ce qu'on peut
      results.push({
        id: String(row[0] || 'ERREUR-LIGNE-' + (i+2)),
        dateCreation: '',
        derniereModif: '',
        artiste: { nom: String(row[3] || '(erreur lecture)'), adresse: '', siret: '' },
        expo: { debut: '', fin: '', lieu: '' },
        installation: { date: '', heure: '' },
        vernissage: { date: '', heure: '' },
        demontage: { date: '', heure: '' },
        droits: { montant: '', tva: '' },
        transport: { actif: '', details: '' },
        hebergement: { actif: '', details: '' },
        actionsCulturelles: { actif: '', details: '' },
        articlesPerso: '[]',
        articlesStd: '{}',
        dateSignature: '',
        statut: 'Erreur',
        lienPdf: ''
      });
    }
  }
  
  results.reverse(); // Plus récentes en premier
  return results;
}

// === FORMATER UNE DATE DE MANIÈRE SÉCURISÉE ===
// Avec heure (pour dateCreation, derniereModif)
function formatDateSafe_(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, 'Europe/Paris', 'dd/MM/yyyy HH:mm');
    }
    var d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, 'Europe/Paris', 'dd/MM/yyyy HH:mm');
  } catch(e) {
    return String(val);
  }
}

// Sans heure, format long français (pour dates d'expo, installation, vernissage, démontage, signature)
// Ex: "lundi 9 février 2026"
function formatDateOnly_(val) {
  if (!val) return '';
  try {
    var d;
    if (val instanceof Date) {
      d = val;
    } else {
      var s = String(val);
      // Format dd/MM/yyyy ? Parser manuellement
      var parts = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (parts) {
        d = new Date(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1]));
      } else {
        d = new Date(val);
      }
    }
    if (isNaN(d.getTime())) return String(val);
    
    var jours = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
    var mois = ['janvier', 'f\u00e9vrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'ao\u00fbt', 'septembre', 'octobre', 'novembre', 'd\u00e9cembre'];
    
    return jours[d.getDay()] + ' ' + d.getDate() + ' ' + mois[d.getMonth()] + ' ' + d.getFullYear();
  } catch(e) {
    return String(val);
  }
}

// Heure seule (pour les champs heure)
function formatTimeSafe_(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, 'Europe/Paris', 'HH:mm');
    }
    var s = String(val);
    // Déjà au format HH:mm ? On le garde
    if (/^\d{1,2}:\d{2}$/.test(s)) return s;
    // Tenter de parser
    var d = new Date(val);
    if (isNaN(d.getTime())) return s;
    return Utilities.formatDate(d, 'Europe/Paris', 'HH:mm');
  } catch(e) {
    return String(val);
  }
}

// === RÉCUPÉRER UNE CONVENTION PAR ID ===
function getConvention(id) {
  const conventions = getConventions();
  return conventions.find(function(c) { return c.id === id; }) || null;
}

// === ENREGISTRER UNE CONVENTION (nouvelle ou modification) ===
function saveConvention(data) {
  const sheet = getSheet_();
  const now = new Date();
  const isNew = !data.id;
  const id = isNew ? generateId_() : data.id;

  var a2 = data.artiste2 || {};
  var a2Actif = (a2.actif === true || a2.actif === 'true' || a2.actif === 'Oui') && !!(a2.nom);

  var rowData = [
    id,
    isNew ? now : '', // sera géré différemment pour modif
    now,
    data.artiste.nom,
    data.artiste.adresse,
    data.artiste.siret,
    a2Actif ? 'Oui' : 'Non',
    a2Actif ? (a2.nom || '') : '',
    a2Actif ? (a2.adresse || '') : '',
    a2Actif ? (a2.siret || '') : '',
    data.expo.debut,
    data.expo.fin,
    data.expo.lieu || 'Espace Culturel L\'Hermine',
    data.installation.date,
    data.installation.heure,
    data.vernissage.date,
    data.vernissage.heure,
    data.demontage.date,
    data.demontage.heure,
    data.droits.montant,
    data.droits.tva,
    data.transport.actif ? 'Oui' : 'Non',
    data.transport.details || '',
    data.hebergement.actif ? 'Oui' : 'Non',
    data.hebergement.details || '',
    data.actionsCulturelles.actif ? 'Oui' : 'Non',
    data.actionsCulturelles.details || '',
    data.dateSignature,
    isNew ? (data.statut || 'Brouillon') : (data.statut || null),
    '', // Lien PDF
    JSON.stringify(data.articlesPerso || []),
    JSON.stringify(data.articlesStd || {})
  ];

  if (isNew) {
    sheet.appendRow(rowData);
  } else {
    // Chercher la ligne existante
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'Aucune convention à modifier' };
    const allData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var found = false;
    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][0]) === String(id)) {
        // Conserver la date de création d'origine
        const dateCreation = sheet.getRange(i + 2, 2).getValue();
        rowData[1] = dateCreation;
        // Conserver aussi le lien PDF existant (colonne 30 = index 29)
        var existingLink = sheet.getRange(i + 2, 30).getValue();
        if (existingLink) rowData[29] = existingLink;
        // Conserver le statut existant si non spécifié (colonne 29 = index 28)
        if (!rowData[28]) {
          var existingStatut = sheet.getRange(i + 2, 29).getValue();
          rowData[28] = existingStatut || 'Brouillon';
        }
        sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'Convention non trouvée pour modification' };
  }

  SpreadsheetApp.flush();
  return { success: true, id: id };
}

// === SUPPRIMER UNE CONVENTION ===
function deleteConvention(id) {
  const sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'Aucune convention trouvée' };
  const allData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false, error: 'Convention non trouvée' };
}

// === SCANNER LE DOSSIER DRIVE ET IMPORTER LES CONVENTIONS MANQUANTES ===
function scanDriveFolder() {
  var folder;
  var folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (!folders.hasNext()) {
    return { success: true, imported: 0, message: 'Dossier Drive non trouvé' };
  }
  folder = folders.next();

  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();

  // Récupérer tous les liens PDF/HTML déjà enregistrés dans le sheet
  var existingLinks = {};
  var existingNames = {};
  if (lastRow > 1) {
    var linkCol = sheet.getRange(2, 30, lastRow - 1, 1).getValues(); // Colonne "Lien PDF"
    var nameCol = sheet.getRange(2, 4, lastRow - 1, 1).getValues();  // Colonne "Artiste - Nom"
    for (var i = 0; i < linkCol.length; i++) {
      if (linkCol[i][0]) existingLinks[linkCol[i][0]] = true;
      if (nameCol[i][0]) existingNames[nameCol[i][0].toLowerCase().trim()] = true;
    }
  }

  // Scanner les fichiers HTML du dossier
  var files = folder.getFilesByType('text/html');
  var imported = 0;
  var skipped = 0;
  var details = [];

  while (files.hasNext()) {
    var file = files.next();
    var fileUrl = file.getUrl();
    var fileName = file.getName();

    // Vérifier si ce fichier est déjà lié dans le sheet
    if (existingLinks[fileUrl]) {
      skipped++;
      continue;
    }

    // Essayer d'extraire des infos du nom de fichier
    // Format attendu : Convention_Prenom_Nom.html
    var artisteName = fileName
      .replace(/^Convention_/i, '')
      .replace(/\.html$/i, '')
      .replace(/_/g, ' ')
      .trim();

    // Vérifier si un artiste avec ce nom existe déjà (éviter les doublons)
    if (existingNames[artisteName.toLowerCase()]) {
      // L'artiste existe mais sans lien → mettre à jour le lien
      if (lastRow > 1) {
        var allNames = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
        var allLinks = sheet.getRange(2, 30, lastRow - 1, 1).getValues();
        for (var j = 0; j < allNames.length; j++) {
          if (allNames[j][0] && allNames[j][0].toLowerCase().trim() === artisteName.toLowerCase() && !allLinks[j][0]) {
            sheet.getRange(j + 2, 30).setValue(fileUrl);
            sheet.getRange(j + 2, 29).setValue('Généré');
            details.push('Lien mis à jour : ' + artisteName);
            imported++;
            break;
          }
        }
      }
      continue;
    }

    // Essayer de lire le contenu HTML pour en extraire des infos
    var parsedData = parseHtmlConvention_(file);

    var now = new Date();
    var id = 'CONV-IMP-' + Math.random().toString(36).substring(2, 8).toUpperCase();

    var rowData = [
      id,
      file.getDateCreated(),
      now,
      parsedData.artisteNom || artisteName || fileName,
      parsedData.artisteAdresse || '',
      parsedData.artisteSiret || '',
      parsedData.expoDebut || '',
      parsedData.expoFin || '',
      parsedData.expoLieu || 'Espace Culturel L\'Hermine',
      parsedData.installDate || '',
      parsedData.installHeure || '',
      parsedData.vernissageDate || '',
      parsedData.vernissageHeure || '',
      parsedData.demontageDate || '',
      parsedData.demontageHeure || '',
      parsedData.droitsMontant || '',
      parsedData.droitsTva || '',
      'Non', '', // Transport
      'Non', '', // Hébergement
      'Non', '', // Actions culturelles
      parsedData.dateSignature || '',
      'Généré (importé)',
      fileUrl
    ];

    sheet.appendRow(rowData);
    imported++;
    details.push('Importé : ' + (parsedData.artisteNom || artisteName));
  }

  SpreadsheetApp.flush();
  return {
    success: true,
    imported: imported,
    skipped: skipped,
    details: details,
    message: imported > 0
      ? imported + ' convention(s) importée(s) depuis le Drive'
      : 'Aucune nouvelle convention trouvée sur le Drive'
  };
}

// === PARSER UN FICHIER HTML DE CONVENTION POUR EN EXTRAIRE LES INFOS ===
function parseHtmlConvention_(file) {
  var result = {
    artisteNom: '', artisteAdresse: '', artisteSiret: '',
    expoDebut: '', expoFin: '', expoLieu: '',
    installDate: '', installHeure: '',
    vernissageDate: '', vernissageHeure: '',
    demontageDate: '', demontageHeure: '',
    droitsMontant: '', droitsTva: '', dateSignature: ''
  };

  try {
    var content = file.getBlob().getDataAsString();

    // Extraire le nom de l'artiste (premier party-block > party-name)
    var nameMatch = content.match(/party-name[^>]*>([^<]+)/);
    if (nameMatch) result.artisteNom = nameMatch[1].trim();

    // SIRET
    var siretMatch = content.match(/SIRET\s*N[°o]\s*(\d[\d\s]*\d)/i);
    if (siretMatch) result.artisteSiret = siretMatch[1].trim();

    // Dates d'exposition (du ... au ...)
    var datesMatch = content.match(/du\s+(\d{1,2}[\/\s]\w+[\/\s]\d{2,4})\s+au\s+(\d{1,2}[\/\s]\w+[\/\s]\d{2,4})/i);
    if (!datesMatch) datesMatch = content.match(/du\s+(\d{1,2}\s+\w+\s+\d{4})\s+au\s+(\d{1,2}\s+\w+\s+\d{4})/i);
    if (!datesMatch) datesMatch = content.match(/du\s+([\d\/]+)\s+au\s+([\d\/]+)/i);
    if (datesMatch) {
      result.expoDebut = datesMatch[1].trim();
      result.expoFin = datesMatch[2].trim();
    }

    // Montant des droits
    var montantMatch = content.match(/somme\s+de\s+(?:<[^>]*>)*\s*(\d[\d\s,.]*)\s*euros/i);
    if (montantMatch) result.droitsMontant = montantMatch[1].replace(/\s/g, '').replace(',', '.').trim();

    // Date de signature
    var sigMatch = content.match(/Fait\s+[àa]\s+Sarzeau\s+en\s+\d+\s+exemplaires?,\s+le\s+([^.<]+)/i);
    if (sigMatch) result.dateSignature = sigMatch[1].trim();

  } catch (e) {
    Logger.log('Erreur parsing HTML : ' + e.message);
  }

  return result;
}

// === GÉNÉRER LE HTML DE LA CONVENTION ===
function generateConventionHtml(data, logos) {
  // Logos en base64 (ou fallback sur le nom de fichier)
  logos = logos || {};
  var logo1Src = logos.logo1 || 'logo.jpg';
  var logo2Src = logos.logo2 || 'logo2.jpg';
  // Numérotation dynamique des articles
  var articleNum = 0;
  function nextArticle() { articleNum++; return articleNum; }
  var articlesStd = data && data.articlesStd ? data.articlesStd : {};

  var html = '<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Convention d\u2019exposition \u2014 ' + escapeHtml_(data.artiste.nom) + '</title>';
  html += getConventionCSS_();
  html += '</head><body>';

  // === PAGE 1 ===
  html += '<header class="header-page1">';
  html += '<div class="logo-container">';
  html += '<img src="' + logo1Src + '" alt="Logo Mairie de Sarzeau" class="logo-page-1">';
  html += '</div>';
  html += '<div class="header-info">';
  html += '<div class="mairie-title">Mairie de Sarzeau</div>';
  html += '<div class="mairie-details">';
  html += 'Place Richemont - BP 14<br>';
  html += '56370 Sarzeau<br>';
  html += 'T\u00e9l. : 02 97 41 85 15';
  html += '</div>';
  html += '<div class="mairie-web">www.sarzeau.fr</div>';
  html += '</div></header>';

  html += '<div class="content-wrapper">';
  html += '<div class="main-title"><h1>Convention d\u2019exposition</h1><div class="underline"></div></div>';

  // Parties
  html += '<section class="parties-section">';
  html += '<p class="parties-intro">ENTRE LES SOUSSIGN\u00c9S,</p>';

  var a2 = data.artiste2 || {};
  var a2Actif = a2.actif === true && !!(a2.nom);

  html += '<div class="party-block">';
  html += '<p class="party-name">' + escapeHtml_(data.artiste.nom) + '</p>';
  html += '<p>' + escapeHtml_(data.artiste.adresse) + '</p>';
  html += '<p>Num\u00e9ro SIRET : ' + escapeHtml_(data.artiste.siret) + '</p>';
  if (a2Actif) {
    html += '<p class="designation">Ci-apr\u00e8s d\u00e9sign\u00e9 \u00ab <strong>l\u2019Exposant 1</strong> \u00bb,</p>';
  } else {
    html += '<p class="designation">Ci-apr\u00e8s d\u00e9sign\u00e9 \u00ab <strong>l\u2019Exposant</strong> \u00bb, d\u2019une part</p>';
  }
  html += '</div>';

  if (a2Actif) {
    html += '<p class="separator-et">\u2014 ET \u2014</p>';
    html += '<div class="party-block">';
    html += '<p class="party-name">' + escapeHtml_(a2.nom) + '</p>';
    if (a2.adresse) html += '<p>' + escapeHtml_(a2.adresse) + '</p>';
    if (a2.siret) html += '<p>Num\u00e9ro SIRET : ' + escapeHtml_(a2.siret) + '</p>';
    html += '<p class="designation">Ci-apr\u00e8s d\u00e9sign\u00e9 \u00ab <strong>l\u2019Exposant 2</strong> \u00bb, d\u2019une part</p>';
    html += '</div>';
  }

  html += '<p class="separator-et">\u2014 ET \u2014</p>';

  html += '<div class="party-block">';
  html += '<p class="party-name">Commune de Sarzeau \u2014 Espace Culturel l\u2019Hermine</p>';
  html += '<p>Place Richemont - BP 14 - 56370 SARZEAU</p>';
  html += '<p>T\u00e9l. : 02 97 48 29 40 \u2014 E.mail : lhermine@sarzeau.fr</p>';
  html += '<p>Num\u00e9ro SIRET : 215 602 400 00016 \u2014 Code NAF/APE : 8411Z</p>';
  html += '<p>Licence d\u2019entrepreneur de spectacles n\u00b0 : PLATESV-D-2022-008045</p>';
  html += '<p>Repr\u00e9sent\u00e9e par : Monsieur Jean-Marc Dupeyrat en sa qualit\u00e9 de Maire</p>';
  html += '<p class="designation">Ci-apr\u00e8s d\u00e9sign\u00e9e \u00ab <strong>l\u2019Organisateur</strong> \u00bb, d\u2019autre part.</p>';
  html += '</div></section></div>';

  html += '<div class="page-break"></div>';

  // === PAGES 2+ : ARTICLES avec header qui se répète ===
  html += '<table class="articles-table"><thead><tr><td>';
  html += '<div class="header-continuation">';
  html += '<img src="' + logo2Src + '" alt="Logo Sarzeau" class="logo-small">';
  html += '<span class="page-title">Convention d\u2019exposition \u2014 ' + escapeHtml_(data.artiste.nom) + '</span>';
  html += '</div>';
  html += '</td></tr></thead>';
  html += '<tbody><tr><td>';
  html += '<div class="content-wrapper">';
  html += '<div class="agreement-clause">IL A \u00c9T\u00c9 ARR\u00caT\u00c9 ET CONVENU CE QUI SUIT :</div>';

  // Article 1 — Objet
  var n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' \u2014 Objet</h3>';
  var customText = getStdOverrideText_(articlesStd, 'objet');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>La présente convention a pour but de définir les modalités d’accueil de l’exposition des œuvres de l’Exposant.</p>';
    html += '<p><span class="highlight">Lieu :</span> ' + escapeHtml_(data.expo.lieu || 'Espace Culturel L’Hermine') + '</p>';
    html += '<p><span class="highlight">Dates de l’exposition :</span> du ' + escapeHtml_(data.expo.debut) + ' au ' + escapeHtml_(data.expo.fin) + '</p>';
  }
  html += '</article>';

  // Article 2 — Obligations de l'Exposant
  n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' \u2014 Obligations de l\u2019Exposant</h3>';
  var customText = getStdOverrideText_(articlesStd, 'obligationsExposant');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>L’Exposant s’engage à fournir l’exposition conformément à ce qui avait été décidé avec l’Organisateur.</p>';
    html += '<p>L’Exposant en outre s’engage à communiquer à l’organisateur une estimation de la valeur d’assurance (clou à clou) de l’exposition. En cas de non-respect de cet article, la collectivité ne prendra pas en charge l’assurance des biens exposés.</p>';
  }
  html += '</article>';

  // Article 3 — Obligations de l'Organisateur
  n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' \u2014 Obligations de l\u2019Organisateur</h3>';
  var customText = getStdOverrideText_(articlesStd, 'obligationsOrganisateur');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>L’Organisateur s’engage à faciliter l’accueil de l’exposition par :</p>';
    html += '<ul>';
    html += '<li>La mise à disposition gratuite de locaux adéquats,</li>';
    html += '<li>La mise à disposition gratuite d’une personne pour aider à l’installation et au démontage de l’exposition.</li>';
    html += '</ul>';
  }
  html += '</article>';

  // Article 4 — Durée
  n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' \u2014 Dur\u00e9e</h3>';
  var customText = getStdOverrideText_(articlesStd, 'duree');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>La présente convention est conclue pour toute la durée de la présentation de l’exposition dans les locaux de l’Organisateur, de son installation à son démontage.</p>';
  }
  html += '</article>';

  // Article 5 — Installation, vernissage et démontage
  n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' \u2014 Installation, vernissage et d\u00e9montage</h3>';
  var customText = getStdOverrideText_(articlesStd, 'installation');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>L’Organisateur et l’Exposant installeront l’exposition le <span class="highlight">' + escapeHtml_(data.installation.date) + ' à ' + escapeHtml_(data.installation.heure) + '</span> à l’Hermine à Sarzeau pour un vernissage de l’exposition le <span class="highlight">' + escapeHtml_(data.vernissage.date) + ' à ' + escapeHtml_(data.vernissage.heure) + '</span>.</p>';
    html += '<p>Le démontage étant le dernier jour de la présentation public, soit le <span class="highlight">' + escapeHtml_(data.demontage.date) + ' à ' + escapeHtml_(data.demontage.heure) + '</span>.</p>';
  }
  html += '</article>';
  // Article 6 — Droits de présentation publique
  n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' — Droits de présentation publique</h3>';
  var customText = getStdOverrideText_(articlesStd, 'droits');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>L’Organisateur versera la somme de <span class="highlight">' + escapeHtml_(data.droits.montant) + ' euros TTC</span> à l’Exposant pour les droits de présentation publique de son exposition à l’Hermine.</p>';
    if (data.droits.tva === 'Non assujetti') {
      html += '<p>TVA non applicable conformément à l’article 293B du CGI. ';
    } else if (data.droits.tva) {
      html += '<p>' + escapeHtml_(data.droits.tva) + '. ';
    } else {
      html += '<p>';
    }
    html += 'Le règlement des sommes dues à l’Exposant sera effectué après dépôt d’une facture sur le portail dématérialisé Chorus Pro, par virement bancaire sur le compte de l’Exposant dans les 30 jours suivant la date de fin d’exposition mentionnée à l’article 1 du présent contrat. Aucun acompte ne peut être versé à la signature du présent contrat.</p>';
  }
  html += '</article>';

  // === Articles optionnels (insérés avant Résiliation) ===

  // Transport
  if (data.transport.actif) {
    n = nextArticle();
    html += '<article class="article no-break">';
    html += '<h3>Article ' + n + ' \u2014 Transport des \u0153uvres</h3>';
    html += '<p>' + escapeHtml_(data.transport.details || 'Le transport des \u0153uvres est \u00e0 la charge de l\u2019Organisateur. Les modalit\u00e9s seront d\u00e9finies en concertation entre les deux parties.') + '</p>';
    html += '</article>';
  }

  // Hébergement
  if (data.hebergement.actif) {
    n = nextArticle();
    html += '<article class="article no-break">';
    html += '<h3>Article ' + n + ' \u2014 H\u00e9bergement</h3>';
    html += '<p>' + escapeHtml_(data.hebergement.details || 'L\u2019Organisateur prend en charge l\u2019h\u00e9bergement de l\u2019Exposant pour la dur\u00e9e n\u00e9cessaire \u00e0 l\u2019installation et au vernissage de l\u2019exposition.') + '</p>';
    html += '</article>';
  }

  // Actions culturelles
  if (data.actionsCulturelles.actif) {
    n = nextArticle();
    html += '<article class="article no-break">';
    html += '<h3>Article ' + n + ' \u2014 Actions culturelles</h3>';
    html += '<p>' + escapeHtml_(data.actionsCulturelles.details || 'Des actions de m\u00e9diation culturelle pourront \u00eatre organis\u00e9es autour de l\u2019exposition (rencontres avec le public, ateliers, visites guid\u00e9es), avec ou sans la pr\u00e9sence de l\u2019Exposant. Les modalit\u00e9s seront d\u00e9finies d\u2019un commun accord.') + '</p>';
    html += '</article>';
  }

  // === Articles personnalisés ===
  var articlesPerso = data.articlesPerso || [];
  if (typeof articlesPerso === 'string') {
    try { articlesPerso = JSON.parse(articlesPerso); } catch(e) { articlesPerso = []; }
  }
  if (Array.isArray(articlesPerso)) {
    for (var ap = 0; ap < articlesPerso.length; ap++) {
      var artPerso = articlesPerso[ap];
      if (artPerso.titre || artPerso.contenu) {
        n = nextArticle();
        html += '<article class="article no-break">';
        html += '<h3>Article ' + n + ' \u2014 ' + escapeHtml_(artPerso.titre || 'Article suppl\u00e9mentaire') + '</h3>';
        html += '<p>' + escapeHtml_(artPerso.contenu || '') + '</p>';
        html += '</article>';
      }
    }
  }

  // Article — Résiliation (toujours présent)
  n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' \u2014 R\u00e9siliation</h3>';
  var customText = getStdOverrideText_(articlesStd, 'resiliation');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>Toute rupture de cette convention, hors cas de force majeure, engendrera de la part de la partie défaillante, une compensation estimée sur la base du montant des frais engagés par l’autre partie. Une annulation générée par des mesures COVID sera considérée comme force majeure et n’engendrera pas de compensation pour l’Exposant.</p>';
  }
  html += '</article>';

  // Article — Assurances
  n = nextArticle();
  html += '<article class="article no-break">';
  html += '<h3>Article ' + n + ' \u2014 Assurances</h3>';
  var customText = getStdOverrideText_(articlesStd, 'assurances');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>L’Organisateur déclare avoir souscrit les assurances nécessaires à la couverture des risques liés au transport et à l’exposition des œuvres de l’Exposant.</p>';
  }
  html += '</article>';
  // Article — Attribution de compétences + Signatures
  // Groupés ensemble pour que la signature ne soit jamais seule sur une page
  n = nextArticle();
  html += '<div class="no-break">';
  html += '<article class="article">';
  html += '<h3>Article ' + n + ' — Attribution de compétences</h3>';
  var customText = getStdOverrideText_(articlesStd, 'competences');
  if (customText) {
    html += formatArticleTextToHtml_(customText);
  } else {
    html += '<p>En cas de litige portant sur l’interprétation ou l’application du présent contrat, les parties conviennent de s’en remettre à l’appréciation des tribunaux compétents de la ville de Rennes.</p>';
  }
  html += '</article>';

  // Signatures
  html += '<section class="signatures-section">';
  var nbExemplaires = a2Actif ? '3' : '2';
  html += '<p class="date-line">Fait \u00e0 Sarzeau en ' + nbExemplaires + ' exemplaires, le ' + escapeHtml_(data.dateSignature) + '.</p>';
  html += '<div class="signatures">';
  html += '<div class="signature-box">';
  html += '<div class="title">Pour la Commune de Sarzeau</div>';
  html += '<p class="name">M. Jean-Marc DUPEYRAT, Maire</p>';
  html += '</div>';
  if (a2Actif) {
    html += '<div class="signature-box">';
    html += '<div class="title">L\u2019Exposant 1</div>';
    html += '<p class="name">' + escapeHtml_(data.artiste.nom) + '</p>';
    html += '</div>';
    html += '<div class="signature-box">';
    html += '<div class="title">L\u2019Exposant 2</div>';
    html += '<p class="name">' + escapeHtml_(a2.nom) + '</p>';
    html += '</div>';
  } else {
    html += '<div class="signature-box">';
    html += '<div class="title">L\u2019Exposant</div>';
    html += '<p class="name">' + escapeHtml_(data.artiste.nom) + '</p>';
    html += '</div>';
  }
  html += '</div></section>';
  html += '</div>'; // fin du no-break

  html += '</div>'; // fin content-wrapper
  html += '</td></tr></tbody></table>'; // fin du table articles

  html += '</body></html>';
  return html;
}

// === CSS DE LA CONVENTION ===
function getConventionCSS_() {
  return '<style>' +
    ':root{--bleu-marine:#1e3a5f;--bleu-clair:#2c5282;--gris-texte:#2d3748;--gris-clair:#e2e8f0;--gris-moyen:#718096;--fond-subtil:#f8fafc}' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:"Segoe UI","Helvetica Neue",Arial,sans-serif;line-height:1.6;color:var(--gris-texte);max-width:21cm;margin:0 auto;padding:1.5cm 2cm;background:#fff;font-size:11pt}' +
    '@media print{body{width:21cm;padding:0}@page{size:A4;margin:1.5cm 2cm 2cm 2cm}.page-break{page-break-after:always;height:0;margin:0}.no-break{page-break-inside:avoid}}' +
    /* En-tête page 1 : logo à gauche, infos à droite */
    '.header-page1{margin-bottom:30px;padding:0}' +
    '.logo-container{margin-bottom:15px}' +
    '.logo-page-1{width:66%;max-width:none;height:auto;display:block}' +
    '.header-info{font-size:9.5pt;color:#666;line-height:1.4}' +
    '.mairie-title{font-family:Georgia,\"Times New Roman\",serif;font-style:italic;font-size:13pt;color:var(--bleu-marine);margin-bottom:6px}' +
    '.mairie-details{font-size:9pt;color:#666;margin-bottom:2px}' +
    '.mairie-web{font-size:9pt;color:var(--bleu-marine);font-style:italic}' +
    /* En-tête pages suivantes */
    '.header-continuation{display:flex;justify-content:space-between;align-items:center;padding-bottom:15px;margin-bottom:25px;border-bottom:2px solid var(--gris-clair)}' +
    '.articles-table{width:100%;border-collapse:collapse;border:none}' +
    '.articles-table td{padding:0;border:none;vertical-align:top}' +
    '.articles-table thead{display:table-header-group}' +
    '.articles-table thead td{padding-bottom:0}' +
    '.logo-small{max-width:120px;height:auto}' +
    '.header-continuation .page-title{font-size:10pt;color:var(--gris-moyen);font-style:italic}' +
    /* Titre principal */
    '.content-wrapper{padding:0 10px}' +
    '.main-title{text-align:center;margin:30px 0 50px 0}' +
    '.main-title h1{font-size:22pt;font-weight:300;color:var(--bleu-marine);text-transform:uppercase;letter-spacing:4px;margin-bottom:10px}' +
    '.main-title .underline{width:80px;height:3px;background:var(--bleu-clair);margin:0 auto}' +
    /* Parties */
    '.parties-section{margin-bottom:40px}' +
    '.parties-intro{font-weight:600;color:var(--bleu-marine);margin-bottom:20px;font-size:11pt}' +
    '.party-block{background:var(--fond-subtil);padding:20px 25px;margin-bottom:20px;border-left:4px solid var(--bleu-clair)}' +
    '.party-block .party-name{font-weight:600;font-size:12pt;color:var(--bleu-marine);margin-bottom:8px}' +
    '.party-block p{margin:4px 0;font-size:10.5pt}' +
    '.party-block .designation{margin-top:12px;font-style:italic;color:var(--gris-moyen)}' +
    '.separator-et{text-align:center;font-style:italic;color:var(--gris-moyen);margin:25px 0;font-size:12pt}' +
    /* Clause d accord */
    '.agreement-clause{text-align:center;font-weight:600;font-size:12pt;color:var(--bleu-marine);margin:40px 0;padding:15px;border-top:1px solid var(--gris-clair);border-bottom:1px solid var(--gris-clair)}' +
    /* Articles */
    '.article{margin-bottom:25px}' +
    '.article h3{font-size:11pt;font-weight:600;color:var(--bleu-marine);text-transform:uppercase;letter-spacing:1px;margin-bottom:12px;padding-bottom:5px;border-bottom:1px solid var(--gris-clair)}' +
    '.article p{margin-bottom:10px;text-align:justify}' +
    '.article ul{margin:10px 0 15px 25px;list-style-type:none}' +
    '.article ul li{position:relative;padding-left:20px;margin-bottom:8px}' +
    '.article ul li::before{content:"\\2014";position:absolute;left:0;color:var(--bleu-clair);font-weight:bold}' +
    '.article .highlight{font-weight:600;color:var(--bleu-marine)}' +
    /* Signatures */
    '.signatures-section{margin-top:50px}' +
    '.date-line{text-align:right;font-style:italic;margin-bottom:40px;color:var(--gris-moyen)}' +
    '.signatures{display:flex;justify-content:space-between;gap:20px}' +
    '.signature-box{flex:1;min-width:0}' +
    '.signature-box .title{font-size:9pt;font-weight:600;text-transform:uppercase;letter-spacing:1px;color:var(--bleu-marine);padding-bottom:8px;border-bottom:2px solid var(--bleu-clair);margin-bottom:15px}' +
    '.signature-box .name{font-size:10.5pt;margin-bottom:80px}' +
    /* Utilitaires */
    '.spacer-small{height:20px}.spacer-medium{height:40px}' +
  '</style>';
}

// === GÉNÉRER LE PDF ET SAUVER SUR DRIVE ===
function generatePdf(id) {
  // Récupérer les données depuis le sheet
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'Aucune convention trouvée' };

  var numCols = Math.max(sheet.getLastColumn(), CONFIG.HEADERS.length);
  var allData;
  try {
    allData = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  } catch(e) {
    allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  }
  
  var rowIndex = -1;
  for (var i = 0; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) return { success: false, error: 'Convention non trouvée (ID: ' + id + ')' };

  var row = allData[rowIndex];
  var a2Actif = String(row[6] || '') === 'Oui';
  var data = {
    artiste: { nom: String(row[3] || ''), adresse: String(row[4] || ''), siret: String(row[5] || '') },
    artiste2: {
      actif: a2Actif,
      nom: String(row[7] || ''),
      adresse: String(row[8] || ''),
      siret: String(row[9] || '')
    },
    expo: { debut: formatDateOnly_(row[10]), fin: formatDateOnly_(row[11]), lieu: String(row[12] || '') || 'Espace Culturel L\'Hermine' },
    installation: { date: formatDateOnly_(row[13]), heure: formatTimeSafe_(row[14]) },
    vernissage: { date: formatDateOnly_(row[15]), heure: formatTimeSafe_(row[16]) },
    demontage: { date: formatDateOnly_(row[17]), heure: formatTimeSafe_(row[18]) },
    droits: { montant: String(row[19] || ''), tva: String(row[20] || '') },
    transport: { actif: String(row[21]) === 'Oui', details: String(row[22] || '') },
    hebergement: { actif: String(row[23]) === 'Oui', details: String(row[24] || '') },
    actionsCulturelles: { actif: String(row[25]) === 'Oui', details: String(row[26] || '') },
    dateSignature: formatDateOnly_(row[27]),
    articlesPerso: row[30] ? parseJsonSafe_(row[30]) : [],
    articlesStd: row[31] ? parseJsonObjectSafe_(row[31]) : {}
  };

  var logos = getLogosBase64_();
  var htmlContent = generateConventionHtml(data, logos);

  var folder = getOrCreateDriveFolder_();
  var nomFichier = 'Convention_' + data.artiste.nom.replace(/[^a-zA-Z0-9àâäéèêëïîôùûüÿçÀÂÄÉÈÊËÏÎÔÙÛÜŸÇ\s-]/g, '').replace(/\s+/g, '_');

  // Sauvegarder le HTML
  var htmlBlob = Utilities.newBlob(htmlContent, 'text/html', nomFichier + '.html');
  var htmlFile = folder.createFile(htmlBlob);
  // Rendre le fichier HTML accessible par lien (lecture seule)
  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Mettre à jour le lien HTML dans le sheet (colonne 30)
  sheet.getRange(rowIndex + 2, 30).setValue(htmlFile.getUrl());
  // Mettre à jour le statut (colonne 29)
  sheet.getRange(rowIndex + 2, 29).setValue('Généré');
  SpreadsheetApp.flush();

  // URL directe pour ouvrir le HTML dans le navigateur
  var htmlDirectUrl = 'https://drive.google.com/uc?export=view&id=' + htmlFile.getId();

  return {
    success: true,
    htmlUrl: htmlFile.getUrl(),
    htmlDirectUrl: htmlDirectUrl,
    htmlFileId: htmlFile.getId(),
    fileName: nomFichier
  };
}

// === GÉNÉRER LE PDF UNIQUEMENT ===
function generatePdfFile(id) {
  // D'abord générer/mettre à jour le HTML
  var htmlResult = generatePdf(id);
  if (!htmlResult.success) return htmlResult;

  try {
    var folder = getOrCreateDriveFolder_();
    var htmlFile = DriveApp.getFileById(htmlResult.htmlFileId);
    var htmlContent = htmlFile.getBlob().getDataAsString();
    
    // Convertir HTML en PDF via Google Docs (méthode fiable)
    var tempDoc = folder.createFile(Utilities.newBlob(htmlContent, 'text/html', 'temp.html'));
    var pdfBlob = tempDoc.getAs('application/pdf');
    pdfBlob.setName(htmlResult.fileName + '.pdf');
    var pdfFile = folder.createFile(pdfBlob);
    
    // Supprimer le fichier temporaire
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
    
    // Rendre le PDF accessible par lien
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      pdfDirectUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
      fileName: htmlResult.fileName + '.pdf'
    };
  } catch(e) {
    return { success: false, error: 'Erreur génération PDF : ' + e.message };
  }
}

// === OBTENIR L'URL DIRECTE DU HTML ===
function getHtmlDirectUrl(id) {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'Aucune convention' };
  
  var allData = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), CONFIG.HEADERS.length)).getValues();
  for (var i = 0; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) {
      var lienDrive = String(allData[i][29] || '');
      if (!lienDrive) return { success: false, error: 'Aucun fichier généré' };
      // Extraire l'ID du fichier depuis l'URL Drive
      var match = lienDrive.match(/[-\w]{25,}/);
      if (match) {
        var fileId = match[0];
        try {
          var file = DriveApp.getFileById(fileId);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          return {
            success: true,
            directUrl: 'https://drive.google.com/uc?export=view&id=' + fileId
          };
        } catch(e) {
          return { success: false, error: 'Fichier introuvable sur le Drive' };
        }
      }
      return { success: false, error: 'Lien invalide' };
    }
  }
  return { success: false, error: 'Convention non trouvée' };
}

// === PRÉVISUALISER LE HTML ===
function previewConvention(data) {
  var logos = getLogosBase64_();
  return generateConventionHtml(data, logos);
}

// === CHARGER LES LOGOS DEPUIS LE DOSSIER DRIVE ===
function getLogosBase64_() {
  var logos = { logo1: '', logo2: '' };
  
  try {
    var folder = getOrCreateDriveFolder_();
    
    // Chercher logo.jpg ou logo.png (page 1)
    var logoNames = ['logo.jpg', 'logo.jpeg', 'logo.png'];
    for (var i = 0; i < logoNames.length; i++) {
      var files = folder.getFilesByName(logoNames[i]);
      if (files.hasNext()) {
        var file = files.next();
        var blob = file.getBlob();
        var base64 = Utilities.base64Encode(blob.getBytes());
        var mimeType = blob.getContentType();
        logos.logo1 = 'data:' + mimeType + ';base64,' + base64;
        break;
      }
    }
    
    // Chercher logo2.jpg ou logo2.png (pages suivantes)
    var logo2Names = ['logo2.jpg', 'logo2.jpeg', 'logo2.png'];
    for (var j = 0; j < logo2Names.length; j++) {
      var files2 = folder.getFilesByName(logo2Names[j]);
      if (files2.hasNext()) {
        var file2 = files2.next();
        var blob2 = file2.getBlob();
        var base642 = Utilities.base64Encode(blob2.getBytes());
        var mimeType2 = blob2.getContentType();
        logos.logo2 = 'data:' + mimeType2 + ';base64,' + base642;
        break;
      }
    }
  } catch(e) {
    // Pas grave si les logos ne sont pas trouvés
  }
  
  return logos;
}

// === UTILITAIRE ÉCHAPPEMENT HTML ===
function escapeHtml_(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function formatArticleTextToHtml_(text) {
  if (!text) return '';
  var s = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  if (!s) return '';
  var lines = s.split('\n');
  var out = [];
  var para = [];
  var list = [];

  function flushPara() {
    if (!para.length) return;
    out.push('<p>' + escapeHtml_(para.join('\n')).replace(/\n/g, '<br>') + '</p>');
    para = [];
  }
  function flushList() {
    if (!list.length) return;
    out.push('<ul>' + list.map(function(it){ return '<li>' + escapeHtml_(it) + '</li>'; }).join('') + '</ul>');
    list = [];
  }

  for (var i = 0; i < lines.length; i++) {
    var t = String(lines[i] || '');
    var trimmed = t.trim();
    if (!trimmed) {
      flushPara();
      flushList();
      continue;
    }

    var isBullet = false;
    var item = '';
    if (trimmed.indexOf('- ') === 0) { isBullet = true; item = trimmed.substring(2).trim(); }
    else if (trimmed.indexOf('• ') === 0) { isBullet = true; item = trimmed.substring(2).trim(); }
    else if (trimmed.indexOf('* ') === 0) { isBullet = true; item = trimmed.substring(2).trim(); }

    if (isBullet) {
      flushPara();
      list.push(item);
    } else {
      flushList();
      para.push(t);
    }
  }

  flushPara();
  flushList();
  return out.join('');
}

function getStdOverrideText_(articlesStd, key) {
  if (!articlesStd) return '';
  var map = articlesStd;
  if (typeof map === 'string') map = parseJsonObjectSafe_(map);
  if (!map || typeof map !== 'object') return '';
  var item = map[key];
  if (!item) return '';
  if (typeof item === 'string') return String(item).trim();
  if (typeof item === 'object') {
    if (item.actif !== true) return '';
    return String(item.texte || '').trim();
  }
  return '';
}


// === UTILITAIRE PARSE JSON SÉCURISÉ ===
function parseJsonSafe_(val) {
  if (!val) return [];
  try {
    var result = JSON.parse(String(val));
    return Array.isArray(result) ? result : [];
  } catch(e) {
    return [];
  }
}

function parseJsonObjectSafe_(val) {
  if (!val) return {};
  try {
    var result = JSON.parse(String(val));
    return (result && typeof result === 'object' && !Array.isArray(result)) ? result : {};
  } catch(e) {
    return {};
  }
}


// === VÉRIFIER LA SYNC DRIVE / SHEET ===
function checkDriveSync() {
  try {
    var sheet = getSheet_();
    var folder = getOrCreateDriveFolder_();
    var lastRow = sheet.getLastRow();
    var conventionsCount = lastRow > 1 ? lastRow - 1 : 0;
    
    // Compter les fichiers HTML dans le dossier Drive
    var files = folder.getFilesByType('text/html');
    var driveCount = 0;
    while (files.hasNext()) { files.next(); driveCount++; }
    
    return {
      success: true,
      sheetCount: conventionsCount,
      driveCount: driveCount,
      sheetName: sheet.getName(),
      folderUrl: folder.getUrl(),
      sheetUrl: getOrCreateSpreadsheet_().getUrl()
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// === DIAGNOSTIC : comprendre pourquoi les données ne chargent pas ===
function debugSheet() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var sheetNames = sheets.map(function(s) { return s.getName(); });
    
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    var result = {
      ssName: ss.getName(),
      ssId: ss.getId(),
      sheetsAvailable: sheetNames,
      activeSheet: sheet.getName(),
      lastRow: lastRow,
      lastCol: lastCol
    };
    
    if (lastRow >= 2) {
      // Lire la ligne 2 brute
      var row2 = sheet.getRange(2, 1, 1, Math.min(lastCol, 32)).getValues()[0];
      result.row2_raw = row2.map(function(cell, i) {
        return {
          col: i + 1,
          value: String(cell),
          type: typeof cell,
          isDate: cell instanceof Date
        };
      });
    }
    
    return result;
  } catch(e) {
    return { error: e.message, stack: e.stack };
  }
}