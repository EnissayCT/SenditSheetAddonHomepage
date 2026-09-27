/**
 * Sendit Google Sheets Add-on
 * Version: 1.1.0
 * Editor add-on: create and track Sendit parcels from Google Sheets.
 */

var SENDIT_API_BASE_URL = 'https://app.sendit.ma/api/v1';

var SENDIT_STATUS_MAP = {
  'PENDING': 'en attente',
  'TO_PREPARE': 'à préparer',
  'NEW_DESTINATION': 'à changer',
  'TO_PICKUP': 'ramassage en cours',
  'PICKEDUP': 'ramassé',
  'WAREHOUSE': 'entrepôt',
  'TRANSIT': 'en transit',
  'DISTRIBUTED': 'distribué',
  'DELIVERING': 'en cours de livraison',
  'DELIVERED': 'livré',
  'PARTIALLY_DELIVERED': 'livré partiellement',
  'UNREACHABLE': 'injoignable',
  'POSTPONED': 'reporté',
  'SCHEDULED': 'programmé',
  'CANCELED': 'annulé',
  'REJECTED': 'refusé',
  'RETURNED': 'retourné'
};

var SENDIT_STATUS_COLORS = {
  'non traité': '#f4cccc',
  'en attente': '#fff2cc',
  'à préparer': '#fff4e6',
  'à changer': '#e6f2ff',
  'ramassage en cours': '#ffe6f0',
  'ramassé': '#f0e6ff',
  'entrepôt': '#e6e6e6',
  'en transit': '#e6f7ff',
  'distribué': '#fff9e6',
  'en cours de livraison': '#fce5cd',
  'livré': '#d9ead3',
  'livré partiellement': '#c9e4ca',
  'injoignable': '#ffe6cc',
  'reporté': '#e6e6ff',
  'programmé': '#f0f0ff',
  'annulé': '#d0d0d0',
  'refusé': '#ffcccc',
  'retourné': '#cfe2f3'
};

var SENDIT_STATUS_LIST = [
  'non traité', 'en attente', 'à préparer', 'à changer', 'ramassage en cours',
  'ramassé', 'entrepôt', 'en transit', 'distribué', 'en cours de livraison',
  'livré', 'livré partiellement', 'injoignable', 'reporté', 'programmé',
  'annulé', 'refusé', 'retourné'
];

var SENDIT_FINAL_STATUSES = ['livré', 'retourné', 'annulé', 'refusé'];

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Ouvrir Sendit', 'showSidebar')
    .addSeparator()
    .addItem('Configurer les clés API', 'showConfigDialog')
    .addItem('Synchroniser les statuts', 'syncAllDeliveryStatusesMenu')
    .addItem('Rafraîchir les villes', 'refreshDistrictsCacheMenu')
    .addSeparator()
    .addItem('Créer une feuille modèle', 'createTemplateSheetMenu')
    .addItem('Télécharger étiquettes (lignes sélectionnées)', 'downloadLabelsMenu')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showWelcomeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Welcome')
    .setWidth(480)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bienvenue sur Sendit pour Google Sheets');
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Sendit')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Config')
    .setWidth(460)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuration Sendit');
}

function showUserMessage_(title, message) {
  SpreadsheetApp.getUi().alert(title, String(message || ''), SpreadsheetApp.getUi().ButtonSet.OK);
}

function saveApiKeys(publicKey, secretKey, pickupDistrictId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('SENDIT_PUBLIC_KEY', publicKey);
    userProperties.setProperty('SENDIT_SECRET_KEY', secretKey);
    if (pickupDistrictId) {
      userProperties.setProperty('SENDIT_PICKUP_DISTRICT_ID', pickupDistrictId.toString());
    }

    const token = authenticateWithSendit(publicKey, secretKey);
    if (token) {
      return { success: true, message: 'Clés API sauvegardées avec succès.' };
    }
    return { success: false, message: 'Échec de l\'authentification. Vérifiez vos clés.' };
  } catch (error) {
    return { success: false, message: 'Erreur : ' + error.message };
  }
}

function getApiKeys() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    publicKey: userProperties.getProperty('SENDIT_PUBLIC_KEY') || '',
    secretKey: userProperties.getProperty('SENDIT_SECRET_KEY') || '',
    pickupDistrictId: userProperties.getProperty('SENDIT_PICKUP_DISTRICT_ID') || ''
  };
}

function areApiKeysConfigured() {
  const keys = getApiKeys();
  return keys.publicKey !== '' && keys.secretKey !== '';
}

function authenticateWithSendit(publicKey, secretKey) {
  try {
    const keys = getApiKeys();
    const payload = {
      public_key: publicKey || keys.publicKey,
      secret_key: secretKey || keys.secretKey
    };

    const response = UrlFetchApp.fetch(SENDIT_API_BASE_URL + '/login', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());

    if (result.success && result.data && result.data.token) {
      CacheService.getUserCache().put('SENDIT_TOKEN', result.data.token, 3600);
      return result.data.token;
    }
    return null;
  } catch (error) {
    console.error('Erreur d\'authentification:', error);
    return null;
  }
}

function getAuthToken() {
  const cache = CacheService.getUserCache();
  let token = cache.get('SENDIT_TOKEN');
  if (!token) {
    token = authenticateWithSendit();
  }
  return token;
}

function sendApiRequest(endpoint, method, payload) {
  const token = getAuthToken();
  if (!token) {
    throw new Error('Impossible de s\'authentifier. Vérifiez vos clés API.');
  }

  const url = SENDIT_API_BASE_URL + endpoint;
  const options = {
    method: method.toLowerCase(),
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const result = JSON.parse(response.getContentText());

  if (statusCode === 401) {
    CacheService.getUserCache().remove('SENDIT_TOKEN');
    const newToken = authenticateWithSendit();
    if (newToken) {
      options.headers['Authorization'] = 'Bearer ' + newToken;
      const retryResponse = UrlFetchApp.fetch(url, options);
      return JSON.parse(retryResponse.getContentText());
    }
  }

  return result;
}

/**
 * Load districts for the config dialog using keys typed in the form.
 * Does not require keys to already be saved (fixes first-time setup).
 */
function fetchDistrictsForConfig(publicKey, secretKey, forceRefresh) {
  try {
    if (!publicKey || !secretKey) {
      return { success: false, message: 'Entrez d\'abord vos clés API, puis chargez les villes.' };
    }

    const token = authenticateWithSendit(publicKey, secretKey);
    if (!token) {
      return { success: false, message: 'Échec de l\'authentification. Vérifiez vos clés.' };
    }

    if (!forceRefresh) {
      const cached = getDistrictsFromCache();
      if (cached && cached.length > 0) {
        return { success: true, data: cached };
      }
    }

    return fetchAndCacheAllDistricts_();
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getDistricts() {
  try {
    const cachedDistricts = getDistrictsFromCache();
    if (cachedDistricts && cachedDistricts.length > 0) {
      return { success: true, data: cachedDistricts };
    }
    return fetchAndCacheAllDistricts_();
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function fetchAndCacheAllDistricts_() {
  const allDistricts = [];
  let page = 1;
  let hasMore = true;

  while (hasMore) {
    const result = sendApiRequest('/districts?per_page=100&page=' + page, 'GET');
    if (result.success && result.data && result.data.length > 0) {
      allDistricts.push.apply(allDistricts, result.data);
      if (result.next_page_url) {
        page++;
        Utilities.sleep(200);
      } else {
        hasMore = false;
      }
    } else {
      hasMore = false;
    }
  }

  if (allDistricts.length > 0) {
    cacheDistricts(allDistricts);
    return { success: true, data: allDistricts };
  }
  return { success: false, message: 'Impossible de récupérer les villes' };
}

function getDistrictsFromCache() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheet = spreadsheet.getSheetByName('_Sendit_Districts_Cache');
    if (!cacheSheet) {
      return null;
    }

    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 2) {
      return null;
    }

    const data = cacheSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    return data.map(function(row) {
      return {
        id: row[0].toString(),
        ville: row[1],
        name: row[2],
        arabic_name: row[3],
        price: row[4],
        delais: row[5],
        pickup_district: row[6]
      };
    });
  } catch (error) {
    console.error('Error reading districts cache:', error);
    return null;
  }
}

function cacheDistricts(districts) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let cacheSheet = spreadsheet.getSheetByName('_Sendit_Districts_Cache');

    if (!cacheSheet) {
      cacheSheet = spreadsheet.insertSheet('_Sendit_Districts_Cache');
      cacheSheet.hideSheet();
      const headers = ['ID', 'Ville', 'Nom', 'Nom Arabe', 'Prix', 'Délai', 'Pickup District'];
      cacheSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      cacheSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285f4')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    } else {
      const lastRow = cacheSheet.getLastRow();
      if (lastRow > 1) {
        cacheSheet.getRange(2, 1, lastRow - 1, 7).clear();
      }
    }

    const districtData = districts.map(function(d) {
      return [
        d.id,
        d.ville || '',
        d.name || '',
        d.arabic_name || '',
        d.price || '',
        d.delais || '',
        d.pickup_district || 0
      ];
    });

    if (districtData.length > 0) {
      cacheSheet.getRange(2, 1, districtData.length, 7).setValues(districtData);
    }
    return true;
  } catch (error) {
    console.error('Error caching districts:', error);
    return false;
  }
}

function refreshDistrictsCache() {
  try {
    if (!areApiKeysConfigured()) {
      return { success: false, message: 'Veuillez d\'abord configurer vos clés API.' };
    }
    CacheService.getUserCache().remove('SENDIT_TOKEN');
    const token = authenticateWithSendit();
    if (!token) {
      return { success: false, message: 'Échec de l\'authentification. Vérifiez vos clés.' };
    }
    const result = fetchAndCacheAllDistricts_();
    if (result.success) {
      result.message = 'Cache des villes mis à jour (' + result.data.length + ' villes).';
    }
    return result;
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function findDistrictIdByCityName_(cityName, districts) {
  if (!cityName || !districts || districts.length === 0) {
    return '';
  }

  const raw = cityName.toString().trim();
  if (!raw) {
    return '';
  }

  const lower = raw.toLowerCase();
  const withoutPrice = raw.replace(/\s*-\s*[\d.,]+\s*DH\s*$/i, '').trim();

  for (let i = 0; i < districts.length; i++) {
    const d = districts[i];
    const name = (d.name || '').toString().trim();
    const ville = (d.ville || '').toString().trim();
    const priced = name + ' - ' + d.price + ' DH';
    if (name === raw || ville === raw || priced === raw) {
      return d.id;
    }
  }

  for (let j = 0; j < districts.length; j++) {
    const d2 = districts[j];
    const name2 = (d2.name || '').toString().trim();
    const ville2 = (d2.ville || '').toString().trim();
    if (name2.toLowerCase() === lower ||
        ville2.toLowerCase() === lower ||
        name2.toLowerCase() === withoutPrice.toLowerCase() ||
        ville2.toLowerCase() === withoutPrice.toLowerCase()) {
      return d2.id;
    }
  }

  return '';
}

function normalizeMoroccanPhone_(phone) {
  let phoneNumber = String(phone || '').replace(/\s+/g, '').replace(/[^\d]/g, '');
  if (phoneNumber.length === 9 && !phoneNumber.startsWith('0')) {
    phoneNumber = '0' + phoneNumber;
  }
  if (!phoneNumber.startsWith('0')) {
    phoneNumber = '0' + phoneNumber;
  }
  return phoneNumber;
}

function findHeaderIndex_(headers, matcher) {
  for (let i = 0; i < headers.length; i++) {
    if (matcher(headers[i].toString().toLowerCase().trim())) {
      return i;
    }
  }
  return -1;
}

function mapRowToApiData_(headers, rowData, districts) {
  const deliveryData = {};
  headers.forEach(function(header, index) {
    deliveryData[header.toString().toLowerCase().trim()] = rowData[index];
  });

  const cityName = deliveryData['ville'] || deliveryData['city'];
  let districtId = deliveryData['ville_id'] || deliveryData['district_id'];
  if (cityName && !districtId) {
    districtId = findDistrictIdByCityName_(cityName, districts);
  }

  return {
    name: deliveryData['nom'] || deliveryData['name'] || deliveryData['client'],
    phone: deliveryData['téléphone'] || deliveryData['phone'] || deliveryData['tel'],
    address: deliveryData['adresse'] || deliveryData['address'],
    district_id: districtId,
    amount: deliveryData['montant'] || deliveryData['amount'] || deliveryData['prix'],
    comment: deliveryData['commentaire'] || deliveryData['comment'] || deliveryData['note'],
    reference: deliveryData['référence'] || deliveryData['reference'] || deliveryData['ref'],
    allow_open: deliveryData['autoriser_ouverture'] === 'OUI' || deliveryData['allow_open'],
    allow_try: deliveryData['autoriser_essai'] === 'OUI' || deliveryData['allow_try'],
    products: deliveryData['produits'] || deliveryData['products'] || '',
    option_exchange: deliveryData['option_échange'] || deliveryData['option_exchange']
  };
}

function getDeliveryStatuses() {
  try {
    const result = sendApiRequest('/all-status-deliveries', 'GET');
    if (result.success && result.data) {
      return { success: true, data: result.data };
    }
    return { success: false, message: 'Impossible de récupérer les statuts' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function createDelivery(deliveryData) {
  try {
    if (!areApiKeysConfigured()) {
      return { success: false, message: 'Veuillez d\'abord configurer vos clés API.' };
    }

    if (!deliveryData.name || !deliveryData.phone || !deliveryData.address || !deliveryData.district_id) {
      return { success: false, message: 'Les champs nom, téléphone, adresse et ville sont obligatoires.' };
    }

    const apiKeys = getApiKeys();
    const pickupDistrictId = deliveryData.pickup_district_id || apiKeys.pickupDistrictId;
    if (!pickupDistrictId) {
      return { success: false, message: 'Ville de ramassage non configurée. Ouvrez Configuration Sendit.' };
    }

    const phoneNumber = normalizeMoroccanPhone_(deliveryData.phone);
    if (phoneNumber.length !== 10 || !phoneNumber.startsWith('0')) {
      return {
        success: false,
        message: 'Le numéro de téléphone doit commencer par 0 et contenir exactement 10 chiffres. Exemple: 0612345678'
      };
    }

    const payload = {
      pickup_district_id: parseInt(pickupDistrictId, 10),
      district_id: parseInt(deliveryData.district_id, 10),
      name: String(deliveryData.name).trim(),
      phone: phoneNumber,
      address: String(deliveryData.address).trim(),
      amount: parseFloat(deliveryData.amount) || 0,
      comment: String(deliveryData.comment || '').trim(),
      reference: String(deliveryData.reference || '').trim(),
      allow_open: deliveryData.allow_open ? 1 : 0,
      allow_try: deliveryData.allow_try ? 1 : 0,
      products_from_stock: 0,
      products: String(deliveryData.products || '').trim(),
      option_exchange: deliveryData.option_exchange ? 1 : 0
    };

    if (deliveryData.packaging_id) {
      payload.packaging_id = parseInt(deliveryData.packaging_id, 10);
    }

    const result = sendApiRequest('/deliveries', 'POST', payload);

    if (result.success && result.data) {
      return {
        success: true,
        message: 'Colis créé avec succès.',
        data: result.data
      };
    }

    let errorMessage = result.message || 'Erreur lors de la création du colis';
    if (result.data && typeof result.data === 'object') {
      const validationErrors = [];
      for (const field in result.data) {
        if (Array.isArray(result.data[field])) {
          validationErrors.push(field + ': ' + result.data[field].join(', '));
        } else {
          validationErrors.push(field + ': ' + result.data[field]);
        }
      }
      if (validationErrors.length > 0) {
        errorMessage += '\n\nDétails: ' + validationErrors.join('\n');
      }
    }

    return { success: false, message: errorMessage };
  } catch (error) {
    return { success: false, message: 'Erreur : ' + error.message };
  }
}

function createDeliveryFromSelectedRow() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeRange = sheet.getActiveRange();
    const row = activeRange.getRow();

    if (row === 1) {
      return { success: false, message: 'Sélectionnez une ligne de données, pas l\'en-tête.' };
    }

    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const rowData = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

    const statusColIndex = findHeaderIndex_(headers, function(h) {
      return h === 'statut' || h === 'status';
    });

    if (statusColIndex !== -1) {
      const currentStatus = rowData[statusColIndex].toString().toLowerCase();
      if (currentStatus !== 'non traité' && currentStatus !== '') {
        return { success: false, message: 'Cette commande a déjà été traitée. Statut actuel: ' + currentStatus };
      }
    }

    const districts = getDistrictsFromCache() || [];
    const apiData = mapRowToApiData_(headers, rowData, districts);
    if (!apiData.district_id) {
      return {
        success: false,
        message: 'Ville introuvable. Choisissez une ville du modèle ou rafraîchissez les villes.'
      };
    }

    if (statusColIndex !== -1) {
      sheet.getRange(row, statusColIndex + 1).setValue('en attente');
      SpreadsheetApp.flush();
    }

    const result = createDelivery(apiData);
    const codeColIndex = findHeaderIndex_(headers, function(h) {
      return h.indexOf('code') !== -1;
    });

    if (result.success && result.data && result.data.code) {
      if (codeColIndex !== -1) {
        sheet.getRange(row, codeColIndex + 1).setValue(result.data.code);
      }
      if (statusColIndex !== -1) {
        const mappedStatus = mapSenditStatusToSheet(result.data.status);
        const statusCell = sheet.getRange(row, statusColIndex + 1);
        statusCell.setValue(mappedStatus);
        applyStatusColor(statusCell, mappedStatus);
      }
    } else if (statusColIndex !== -1) {
      const statusCell = sheet.getRange(row, statusColIndex + 1);
      statusCell.setValue('non traité');
      applyStatusColor(statusCell, 'non traité');
    }

    return result;
  } catch (error) {
    return { success: false, message: 'Erreur : ' + error.message };
  }
}

function getDeliveryDetails(code) {
  try {
    const result = sendApiRequest('/deliveries/' + code, 'GET');
    if (result.success && result.data) {
      return { success: true, data: result.data };
    }
    return { success: false, message: 'Colis non trouvé' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function mapSenditStatusToSheet(senditStatus) {
  return SENDIT_STATUS_MAP[senditStatus] || 'en attente';
}

function getStatusColor_(status) {
  const statusLower = status.toString().toLowerCase();
  return SENDIT_STATUS_COLORS[statusLower] || '#ffffff';
}

function applyStatusColor(cell, status) {
  cell.setBackground(getStatusColor_(status));
}

function isFinalStatus_(status) {
  return SENDIT_FINAL_STATUSES.indexOf(String(status || '').toLowerCase()) !== -1;
}

function syncSheetStatuses(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn < 1) {
      return { success: true, updatedCount: 0, skippedCount: 0, message: 'Aucune donnée à synchroniser.' };
    }

    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const codeColIndex = findHeaderIndex_(headers, function(h) {
      return h.indexOf('code') !== -1;
    });
    const statusColIndex = findHeaderIndex_(headers, function(h) {
      return h === 'statut' || h === 'status';
    });

    if (codeColIndex === -1) {
      return { success: false, updatedCount: 0, skippedCount: 0, message: 'Colonne "Code" introuvable.' };
    }
    if (statusColIndex === -1) {
      return { success: false, updatedCount: 0, skippedCount: 0, message: 'Colonne "Statut" introuvable.' };
    }

    const numRows = lastRow - 1;
    const data = sheet.getRange(2, 1, numRows, lastColumn).getValues();
    const statusValues = [];
    const statusColors = [];
    let updatedCount = 0;
    let skippedCount = 0;
    let apiCalls = 0;

    for (let i = 0; i < data.length; i++) {
      const code = data[i][codeColIndex];
      const currentStatus = data[i][statusColIndex];
      statusValues.push([currentStatus]);
      statusColors.push([getStatusColor_(currentStatus)]);

      if (isFinalStatus_(currentStatus)) {
        skippedCount++;
        continue;
      }

      if (!code || !code.toString().startsWith('D')) {
        continue;
      }

      const result = getDeliveryDetails(code);
      apiCalls++;
      if (result.success && result.data) {
        const mappedStatus = mapSenditStatusToSheet(result.data.status);
        statusValues[i][0] = mappedStatus;
        statusColors[i][0] = getStatusColor_(mappedStatus);
        updatedCount++;
      }

      if (apiCalls % 10 === 0) {
        Utilities.sleep(500);
      }
    }

    sheet.getRange(2, statusColIndex + 1, numRows, 1).setValues(statusValues);
    sheet.getRange(2, statusColIndex + 1, numRows, 1).setBackgrounds(statusColors);

    return {
      success: true,
      updatedCount: updatedCount,
      skippedCount: skippedCount,
      message: 'Synchronisation terminée : ' + updatedCount + ' colis mis à jour' +
        (skippedCount ? ', ' + skippedCount + ' déjà clôturé(s) ignoré(s)' : '') + '.'
    };
  } catch (error) {
    return { success: false, updatedCount: 0, skippedCount: 0, message: error.message };
  }
}

function syncAllDeliveryStatuses() {
  try {
    if (!areApiKeysConfigured()) {
      return { success: false, message: 'Veuillez d\'abord configurer vos clés API.' };
    }
    return syncSheetStatuses(SpreadsheetApp.getActiveSheet());
  } catch (error) {
    return { success: false, message: 'Erreur lors de la synchronisation : ' + error.message };
  }
}

function syncAllDeliveryStatusesMenu() {
  const result = syncAllDeliveryStatuses();
  showUserMessage_(result.success ? 'Synchronisation' : 'Erreur', result.message);
}

function applyStatusValidations_(sheet, statusColumn, numRows) {
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SENDIT_STATUS_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusColumn, numRows, 1).setDataValidation(statusRule);

  const rules = [];
  const range = sheet.getRange(2, statusColumn, numRows, 1);
  Object.keys(SENDIT_STATUS_COLORS).forEach(function(status) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(SENDIT_STATUS_COLORS[status])
        .setRanges([range])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);
}

function createTemplateSheet() {
  try {
    const districtsResult = getDistricts();
    if (!districtsResult.success || !districtsResult.data) {
      throw new Error('Impossible de charger les villes. Vérifiez votre connexion API.');
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = spreadsheet.insertSheet('Modèle Sendit');
    const headers = [
      'Code', 'Statut', 'Nom', 'Téléphone', 'Adresse',
      'Ville', 'Montant', 'Produits', 'Référence', 'Commentaire',
      'Autoriser_Ouverture', 'Autoriser_Essai'
    ];

    templateSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    templateSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    templateSheet.setColumnWidth(1, 120);
    templateSheet.setColumnWidth(2, 170);
    templateSheet.setColumnWidth(3, 150);
    templateSheet.setColumnWidth(4, 130);
    templateSheet.setColumnWidth(5, 250);
    templateSheet.setColumnWidth(6, 180);
    templateSheet.setColumnWidth(7, 100);
    templateSheet.setColumnWidth(8, 200);
    templateSheet.setColumnWidth(9, 120);
    templateSheet.setColumnWidth(10, 200);
    templateSheet.setColumnWidth(11, 80);
    templateSheet.setColumnWidth(12, 80);

    const cityNames = districtsResult.data
      .map(function(d) { return (d.name || '').toString().trim(); })
      .filter(function(name, index, arr) { return name && arr.indexOf(name) === index; });

    const villeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(cityNames, true)
      .setAllowInvalid(false)
      .setHelpText('Sélectionnez une ville dans la liste')
      .build();
    templateSheet.getRange(2, 6, 1000, 1).setDataValidation(villeRule);

    applyStatusValidations_(templateSheet, 2, 1000);

    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['OUI', 'NON'], true)
      .setAllowInvalid(false)
      .build();
    templateSheet.getRange(2, 11, 1000, 2).setDataValidation(yesNoRule);

    templateSheet.getRange(2, 4, 1000, 1).setNumberFormat('@');

    const exampleRow = [
      '', 'non traité', 'Client Exemple', '0612345678', '123 Rue Example, Casablanca',
      cityNames[0] || '', '199', 'T-Shirt Blanc (x2), Pantalon (x1)', 'REF-001', 'Livraison urgente', 'OUI', 'OUI'
    ];
    templateSheet.getRange(2, 1, 1, exampleRow.length).setValues([exampleRow]);
    applyStatusColor(templateSheet.getRange(2, 2), 'non traité');
    templateSheet.getRange(2, 4).setNumberFormat('@').setValue('0612345678');

    templateSheet.getRange('A1').setNote(
      'Téléphone: saisissez le numéro comme texte (exemple 0612345678). ' +
      'Ville: choisissez le nom exact dans la liste. Statut "non traité" = pas encore envoyé à Sendit.'
    );

    templateSheet.setFrozenRows(1);

    return {
      success: true,
      message: 'Feuille modèle créée. ' + cityNames.length + ' villes disponibles.'
    };
  } catch (error) {
    return { success: false, message: 'Erreur : ' + error.message };
  }
}

function createBulkDeliveriesFromSelection() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeRange = sheet.getActiveRange();
    let startRow = activeRange.getRow();
    let endRow = startRow + activeRange.getNumRows() - 1;

    if (startRow === 1) {
      startRow = 2;
    }
    if (endRow < startRow) {
      return { success: false, message: 'Sélectionnez au moins une ligne de données.' };
    }

    return createBulkDeliveries(startRow, endRow);
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function createBulkDeliveries(startRow, endRow) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const statusColIndex = findHeaderIndex_(headers, function(h) {
      return h === 'statut' || h === 'status';
    });
    const codeColIndex = findHeaderIndex_(headers, function(h) {
      return h.indexOf('code') !== -1;
    });

    const numRows = endRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, lastColumn).getValues();
    const districts = getDistrictsFromCache() || [];

    const results = { success: 0, failed: 0, skipped: 0, errors: [] };
    const statusWrites = [];
    const statusColors = [];
    const codeWrites = [];

    for (let i = 0; i < data.length; i++) {
      statusWrites.push([data[i][statusColIndex] || '']);
      statusColors.push([getStatusColor_(data[i][statusColIndex] || '')]);
      codeWrites.push([codeColIndex !== -1 ? data[i][codeColIndex] : '']);
    }

    for (let i = 0; i < data.length; i++) {
      const rowData = data[i];
      const sheetRow = startRow + i;

      if (statusColIndex !== -1) {
        const currentStatus = rowData[statusColIndex].toString().toLowerCase();
        if (currentStatus !== 'non traité' && currentStatus !== '') {
          results.skipped++;
          continue;
        }
      }

      const apiData = mapRowToApiData_(headers, rowData, districts);
      const result = createDelivery(apiData);

      if (result.success && result.data) {
        results.success++;
        if (codeColIndex !== -1) {
          codeWrites[i][0] = result.data.code;
        }
        if (statusColIndex !== -1) {
          const mappedStatus = mapSenditStatusToSheet(result.data.status);
          statusWrites[i][0] = mappedStatus;
          statusColors[i][0] = getStatusColor_(mappedStatus);
        }
      } else {
        results.failed++;
        results.errors.push('Ligne ' + sheetRow + ': ' + result.message);
        if (statusColIndex !== -1) {
          statusWrites[i][0] = 'non traité';
          statusColors[i][0] = getStatusColor_('non traité');
        }
      }

      Utilities.sleep(300);
    }

    if (codeColIndex !== -1) {
      sheet.getRange(startRow, codeColIndex + 1, numRows, 1).setValues(codeWrites);
    }
    if (statusColIndex !== -1) {
      sheet.getRange(startRow, statusColIndex + 1, numRows, 1).setValues(statusWrites);
      sheet.getRange(startRow, statusColIndex + 1, numRows, 1).setBackgrounds(statusColors);
    }

    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function createTemplateSheetMenu() {
  const result = createTemplateSheet();
  showUserMessage_(result.success ? 'Feuille modèle' : 'Erreur', result.message);
}

function refreshDistrictsCacheMenu() {
  const result = refreshDistrictsCache();
  showUserMessage_(
    result.success ? 'Villes' : 'Erreur',
    result.message || (result.success ? 'Cache mis à jour.' : 'Échec du rafraîchissement.')
  );
}

function downloadLabelsForSelectedRows() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeRange = sheet.getActiveRange();
    const startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();

    if (startRow === 1 && numRows === 1) {
      return { success: false, message: 'Sélectionnez des lignes de données, pas l\'en-tête.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const codeColIndex = findHeaderIndex_(headers, function(h) {
      return h.indexOf('code') !== -1;
    });
    if (codeColIndex === -1) {
      return { success: false, message: 'Colonne "Code" introuvable.' };
    }

    const dataStart = startRow === 1 ? 2 : startRow;
    const dataCount = startRow === 1 ? Math.max(numRows - 1, 0) : numRows;
    if (dataCount < 1) {
      return { success: false, message: 'Sélectionnez des lignes de données, pas l\'en-tête.' };
    }

    const codesRange = sheet.getRange(dataStart, codeColIndex + 1, dataCount, 1).getValues();
    const codes = [];
    for (let i = 0; i < codesRange.length; i++) {
      const code = codesRange[i][0];
      if (code && code.toString().startsWith('D')) {
        codes.push(code.toString());
      }
    }

    if (codes.length === 0) {
      return { success: false, message: 'Aucun code de colis valide trouvé dans la sélection.' };
    }

    return getDeliveryLabels(codes);
  } catch (error) {
    return { success: false, message: 'Erreur : ' + error.message };
  }
}

function getDeliveryLabels(codes) {
  try {
    if (!areApiKeysConfigured()) {
      return { success: false, message: 'Veuillez d\'abord configurer vos clés API.' };
    }

    const token = getAuthToken();
    if (!token) {
      return { success: false, message: 'Impossible de s\'authentifier.' };
    }

    const response = UrlFetchApp.fetch(SENDIT_API_BASE_URL + '/deliveries/getlabels', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        codesToPrint: codes.join(','),
        printFormat: 1
      }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());

    if (result.success && result.data && result.data.fileUrl) {
      return {
        success: true,
        message: 'Étiquettes générées.',
        labelUrl: result.data.fileUrl
      };
    }
    return {
      success: false,
      message: result.message || 'Échec de la génération des étiquettes'
    };
  } catch (error) {
    return { success: false, message: 'Erreur : ' + error.message };
  }
}

function downloadLabelsMenu() {
  const result = downloadLabelsForSelectedRows();
  if (result.success && result.labelUrl) {
    const htmlOutput = HtmlService.createHtmlOutput(
      '<base target="_blank">' +
      '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">' +
      '<div style="padding:16px;font-family:Roboto,Arial,sans-serif;">' +
      '<p>Les étiquettes sont prêtes.</p>' +
      '<p><a href="' + result.labelUrl + '" target="_blank" rel="noopener">Ouvrir le PDF</a></p>' +
      '<button class="action" onclick="google.script.host.close()">Fermer</button>' +
      '</div>'
    ).setWidth(320).setHeight(140);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Étiquettes');
  } else {
    showUserMessage_('Étiquettes', result.message || 'Erreur lors de la génération des étiquettes');
  }
}

function addDeliveryToSheet(apiResponse, deliveryData, cityText) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    let lastColumn = Math.max(sheet.getLastColumn(), 1);
    let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    if (headers.length === 0 || headers[0] === '') {
      const newHeaders = [
        'Code', 'Statut', 'Nom', 'Téléphone', 'Adresse', 'Ville',
        'Montant', 'Produits', 'Référence', 'Commentaire',
        'Autoriser_Ouverture', 'Autoriser_Essai'
      ];
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      sheet.getRange(1, 1, 1, newHeaders.length)
        .setBackground('#4285f4')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      headers = newHeaders;
      lastColumn = newHeaders.length;
    }

    const mappedStatus = mapSenditStatusToSheet(apiResponse.status);
    const rowData = [];
    let phoneCol = -1;

    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().toLowerCase().trim();
      if (header.indexOf('code') !== -1) {
        rowData.push(apiResponse.code || '');
      } else if (header === 'statut' || header === 'status') {
        rowData.push(mappedStatus);
      } else if (header === 'nom' || header === 'name') {
        rowData.push(deliveryData.name || '');
      } else if (header.indexOf('téléphone') !== -1 || header.indexOf('phone') !== -1 || header === 'tel') {
        phoneCol = i;
        rowData.push(normalizeMoroccanPhone_(deliveryData.phone));
      } else if (header.indexOf('adresse') !== -1 || header === 'address') {
        rowData.push(deliveryData.address || '');
      } else if (header === 'ville' || header === 'city') {
        rowData.push(cityText || '');
      } else if (header === 'montant' || header === 'amount' || header === 'prix') {
        rowData.push(deliveryData.amount || '');
      } else if (header === 'produits' || header === 'products') {
        rowData.push(deliveryData.products || '');
      } else if (header.indexOf('référence') !== -1 || header === 'reference' || header === 'ref') {
        rowData.push(deliveryData.reference || '');
      } else if (header === 'commentaire' || header === 'comment' || header === 'note') {
        rowData.push(deliveryData.comment || '');
      } else if (header.indexOf('autoriser_ouverture') !== -1 || header === 'allow_open') {
        rowData.push(deliveryData.allow_open ? 'OUI' : 'NON');
      } else if (header.indexOf('autoriser_essai') !== -1 || header === 'allow_try') {
        rowData.push(deliveryData.allow_try ? 'OUI' : 'NON');
      } else {
        rowData.push('');
      }
    }

    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    if (phoneCol !== -1) {
      sheet.getRange(newRow, phoneCol + 1).setNumberFormat('@').setValue(rowData[phoneCol]);
    }

    const statusColIndex = findHeaderIndex_(headers, function(h) {
      return h === 'statut' || h === 'status';
    });
    if (statusColIndex !== -1) {
      applyStatusColor(sheet.getRange(newRow, statusColIndex + 1), mappedStatus);
    }

    return { success: true };
  } catch (error) {
    console.error('Error adding delivery to sheet:', error);
    return { success: false, message: error.message };
  }
}
