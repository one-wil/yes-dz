// ---------- دوال المساعدة: hashCode ----------
if (typeof String.prototype.hashCode === 'undefined') {
  String.prototype.hashCode = function() {
    var hash = 0, i, chr;
    if (this.length === 0) return hash;
    for (i = 0; i < this.length; i++) {
      chr = this.charCodeAt(i);
      hash = ((hash << 5) - hash) + chr;
      hash |= 0;
    }
    return Math.abs(hash);
  };
}
// ---------- الإعدادات الأساسية ----------
const SPREADSHEET_ID = '1ie0uIyKxLagFof1fhBinOIiCQ9Sy5lRcSYWQfD-GPy0';
const SHEET_NAME = 'Commandes';
const STORE_EMAIL = 'benaakrabahcene@gmail.com';
const STORE_NAME = 'SHOPLIVE';
const ORDERS_API_SECRET_PROPERTY = 'ORDERS_API_SECRET';
const ORDER_STATUSES = ['En attente', 'Confirmée', 'En cours de livraison', 'Livrée', 'Retour'];

// ---------- دالة الدخول الرئيسية ----------
function doPost(e) {
  var requestId = null;
  var lock = LockService.getScriptLock();
  try {
    console.log('=== DÉBUT TRAITEMENT COMMANDE ===');
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Aucun postData fourni');
    }
    const data = JSON.parse(e.postData.contents);

    // كل الاتصالات تمر عبر Worker وتحتاج Secret.
    if (!isValidOrdersApiSecret(data.secret)) {
      return sendResponse('error', 'Accès non autorisé');
    }

    // أوامر الإدارة (list/get/updateStatus/setup) لها مسار API مستقل.
    // Worker هو الذي يحدد StoreKey الموثوق، وليس المتصفح.
    if (data.action && data.action !== 'createOrder') {
      const apiResult = ordersApi(data);
      return ContentService.createTextOutput(JSON.stringify(apiResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // API متعددة المتاجر: الطلب يجب أن يحمل StoreKey الذي حدده Worker.
    const storeKey = normalizeStoreKey(data.storeKey);
    if (!storeKey) {
      return sendResponse('error', 'StoreKey obligatoire');
    }

    const orderData = data.data || data.order || data;
    orderData.storeKey = storeKey;
    if (data.storeName) orderData.storeName = String(data.storeName).trim();

    requestId = generateRequestId(orderData, storeKey);
    console.log('🆔 Request ID:', requestId);

    lock.waitLock(30000);
    console.log('🔒 Lock obtenu');

    if (isRequestProcessed(requestId)) {
      console.log('⚠️ Requête déjà traitée (properties):', requestId);
      lock.releaseLock();
      return sendResponse('success', 'Commande déjà traitée (properties)', { requestId: requestId });
    }

    if (isRequestInSheet(requestId, storeKey)) {
      console.log('⚠️ Requête déjà dans le sheet:', requestId);
      markRequestAsCompleted(requestId);
      lock.releaseLock();
      return sendResponse('success', 'Commande déjà enregistrée dans la feuille', { requestId: requestId });
    }

    markRequestAsProcessing(requestId);
    const orderId = generateOrderId();
    console.log('📋 Order ID généré:', orderId);

    const sheetResult = saveToSpreadsheet(orderData, orderId, requestId, storeKey);
    console.log('💾 Résultat sheet:', sheetResult);

    if (!isEmailSentForRequest(requestId)) {
      console.log('📧 Envoi de l\'email...');
      const emailResult = sendSingleEmail(orderData, orderId, STORE_EMAIL, orderData.storeName || STORE_NAME);
      if (emailResult.success) {
        markEmailAsSent(requestId);
      } else {
        console.warn('⚠️ Probleme envoi email:', emailResult);
      }
      console.log('✅ Résultat email:', emailResult);
    } else {
      console.log('⚠️ Email déjà envoyé pour cette requête');
    }

    markRequestAsCompleted(requestId);
    lock.releaseLock();
    console.log('🔓 Lock released');

    return sendResponse('success', 'Commande envoyée avec succès', {
      orderId: orderId,
      sheetSaved: true,
      requestId: requestId,
      storeKey: storeKey
    });
  } catch (error) {
    console.error('❌ Erreur:', error);
    try { lock.releaseLock(); } catch (e) { }
    if (requestId) cleanupRequest(requestId);
    return sendResponse('error', 'Erreur lors de l\'envoi de la commande: ' + error.toString());
  }
}

// ---------- توليد Request ID ثابت ----------
function generateRequestId(data, storeKey) {
  const customer = (data.customerPhone || '').toString().trim();
  const address = (data.customerAddress || '').toString().trim();
  const itemsString = JSON.stringify(data.items || []);
  const total = (data.total || 0).toString();
  const raw = normalizeStoreKey(storeKey) + '|' + customer + '|' + address + '|' + itemsString + '|' + total;
  return 'REQ_' + raw.hashCode();
}

// ---------- توليد رقم طلب فريد ----------
function generateOrderId() {
  const timestamp = Date.now();
  const random = Math.floor(Math.random() * 10000);
  return 'ORD-' + timestamp + '-' + random;
}

// ---------- التحقق من Script Properties ----------
function isRequestProcessed(requestId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const status = props.getProperty(requestId);
    console.log('🔍 Vérification requête (props):', requestId, '->', status);
    return status === 'DONE' || status === 'PROCESSING';
  } catch (err) {
    console.error('Erreur vérification requête:', err);
    return false;
  }
}
function markRequestAsProcessing(requestId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(requestId, 'PROCESSING');
    console.log('🔄 Marked PROCESSING:', requestId);
  } catch (err) {
    console.error('Erreur marquage traitement:', err);
  }
}
function markRequestAsCompleted(requestId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(requestId, 'DONE');
    console.log('✅ Marked DONE:', requestId);
  } catch (err) {
    console.error('Erreur marquage complétion:', err);
  }
}
function cleanupRequest(requestId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(requestId);
    props.deleteProperty(requestId + '_email');
    console.log('🧹 Requête nettoyée:', requestId);
  } catch (err) {
    console.error('Erreur nettoyage:', err);
  }
}

// ---------- التحقق من وجود requestId داخل الشيت ----------
function isRequestInSheet(requestId, storeKey) {
  try {
    const sheet = getOrdersSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return false;
    const idx = getHeaderIndexes(sheet);
    const values = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, sheet.getLastColumn()).getValues();
    const target = normalizeStoreKey(storeKey);
    return values.some(function(row) {
      return String(row[idx.RequestID] || '').trim() === String(requestId || '').trim() &&
        normalizeStoreKey(row[idx.StoreKey]) === target;
    });
  } catch (err) { console.error('Erreur isRequestInSheet:', err); return false; }
}

// ---------- إدارة الإيميل المرسل ----------
function markEmailAsSent(requestId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(requestId + '_email', 'SENT');
    console.log('📫 Email marqué comme envoyé:', requestId);
  } catch (err) {
    console.error('Erreur marquage email:', err);
  }
}
function isEmailSentForRequest(requestId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const val = props.getProperty(requestId + '_email');
    console.log('🔍 Vérification email:', requestId, '->', val);
    return val === 'SENT';
  } catch (err) {
    console.error('Erreur vérification email:', err);
    return false;
  }
}

// ---------- إرسال إيميل مرة واحدة ----------
function sendSingleEmail(data, orderId, storeEmail, storeName) {
  try {
    const remainingQuota = MailApp.getRemainingDailyQuota();
    console.log('📊 Quota email restant:', remainingQuota);
    if (remainingQuota < 1) return { success: false, message: 'Quota e-mail épuisé' };
    const subject = `🛍️ Nouvelle Commande - ${storeName} - ${orderId}`;
    const body = createEmailBody(data, orderId);
    MailApp.sendEmail({
      to: storeEmail,
      subject: subject,
      htmlBody: body,
      replyTo: data.customerPhone || storeEmail
    });
    console.log('✅ Email envoyé avec succès à:', storeEmail);
    return { success: true, message: 'Email envoyé' };
  } catch (err) {
    console.error('❌ Erreur envoi email:', err);
    return { success: false, message: err.toString() };
  }
}

// ---------- حفظ الطلب في Google Sheet ----------
function saveToSpreadsheet(data, orderId, requestId, storeKey) {
  try {
    const sheet = getOrdersSheet();
    if (isRequestInSheet(requestId, storeKey)) {
      return { success: false, message: 'Duplicate - déjà dans sheet' };
    }

    const idx = getHeaderIndexes(sheet);
    const items = Array.isArray(data.items) ? data.items : [];
    const articlesText = buildArticlesText(items);
    const totalQuantity = items.reduce(function(total, item) {
      return total + (parseInt(item.quantity || 1, 10) || 0);
    }, 0);
    const phone = String(data.customerPhone || '');

    const row = new Array(getOrdersSheetWidth()).fill('');
    row[idx.StoreKey] = storeKey;
    row[idx.Statut] = 'En attente';
    row[idx.RequestID] = requestId;
    row[idx['Date et Heure']] = new Date();
    row[idx['ID Commande']] = orderId;
    row[idx['Nom Client']] = data.customerName || '';
    row[idx['Téléphone']] = phone;
    row[idx.Commune] = data.customerCommune || '';
    row[idx.Adresse] = data.customerAddress || '';
    row[idx.Wilaya] = data.wilaya || '';
    row[idx['Type Livraison']] = data.deliveryType === 'desk' ? 'Bureau' : 'Domicile';
    row[idx['Articles Commandés']] = articlesText;
    row[idx['Quantité Totale']] = totalQuantity;
    row[idx['Sous-total']] = Number(data.subtotal || 0);
    row[idx.Remise] = Number(data.discount || 0);
    row[idx['Frais Livraison']] = Number(data.deliveryCost || 0);
    row[idx['Total Général']] = Number(data.total || 0);
    if (idx.OrderData >= 0) row[idx.OrderData] = JSON.stringify(items);

    const rowNumber = Math.max(sheet.getLastRow() + 1, DATA_START_ROW);
    sheet.getRange(rowNumber, 1, 1, getOrdersSheetWidth()).setValues([row]);
    sheet.getRange(rowNumber, idx['Téléphone'] + 1).setNumberFormat('@').setValue(phone);
    formatOrderRow(sheet, rowNumber);

    return { success: true, row: rowNumber };
  } catch (err) {
    console.error('❌ Erreur sauvegarde Sheet:', err);
    throw err;
  }
}

function buildArticlesText(items) {
  if (!items || items.length === 0) return 'Aucun article';
  return items.map(function(item, index) {
    const quantity = parseInt(item.quantity || 1, 10) || 1;
    const name = item.name || 'Produit';
    const size = item.size || 'N/A';
    const color = item.color || 'N/A';
    const price = Number(item.finalPrice || item.price || 0);
    const total = price * quantity;
    return '🛍️ ' + (index + 1) + '--->: ' + quantity + ' x ' + name + '\n' +
      '   Taille: ' + size + '   |   Couleur: ' + color + '\n' +
      '   Prix: ' + formatNumber(price) + ' DA   |   Total: ' + formatNumber(total) + ' DA';
  }).join('\n\n');
}


// ---------- إنشاء محتوى إيميل ----------
function createEmailBody(data, orderId) {
  const itemsHtml = (data.items || []).map(function(item) {
    const imageHtml = item.image ? `<img src="${item.image}" alt="${item.name || 'Produit'}" style="width:50px;height:50px;object-fit:cover;border-radius:5px;margin-right:8px;">` : '';
    const promoHtml = (item.finalPrice && item.finalPrice !== item.price) ? `<br><small style="color:#FF6B35;">Prix promo: ${formatNumber(item.finalPrice)} DA</small>` : '';
    return `
      <tr>
        <td style="border:1px solid #ddd;padding:8px;">
          <div style="display:flex;align-items:center;">
            ${imageHtml}
            <div>
              <strong>${item.name || 'Non spécifié'}</strong>${promoHtml}
            </div>
          </div>
        </td>
        <td style="border:1px solid #ddd;padding:8px;text-align:center;">${item.quantity || 0}</td>
        <td style="border:1px solid #ddd;padding:8px;text-align:right;">${formatNumber(item.price || 0)} DA</td>
        <td style="border:1px solid #ddd;padding:8px;text-align:right;">${formatNumber((item.finalPrice || item.price || 0) * (item.quantity || 1))} DA</td>
        <td style="border:1px solid #ddd;padding:8px;">${item.size || 'N/A'}</td>
        <td style="border:1px solid #ddd;padding:8px;">
          <div style="display:flex;align-items:center;gap:6px;">
            <div style="width:20px;height:20px;border-radius:50%;background-color:${getColorHex(item.color)};border:1px solid #ddd;"></div>
            ${item.color || 'N/A'}
          </div>
        </td>
      </tr>`;
  }).join('');
  const deliveryType = data.deliveryType === 'desk' ? 'Livraison au bureau' : 'Livraison à domicile';
  const freeDeliveryText = data.isFreeDelivery ? '<span style="color:#06D6A0;font-weight:bold;">OUI (Livraison Gratuite)</span>' : '<span style="color:#FF9E16;">NON</span>';
  return `
  <!DOCTYPE html>
  <html lang="fr">
  <head><meta charset="utf-8"></head>
  <body style="font-family:Arial, sans-serif;color:#333;max-width:800px;margin:0 auto;padding:20px;background:#f9f9f9;">
    <div style="background:linear-gradient(135deg,#4A6CF7,#00D4FF);color:#fff;padding:25px;border-radius:10px;text-align:center;margin-bottom:20px;">
      <h1 style="margin:0;">🎉 NOUVELLE COMMANDE REÇUE 🎉</h1>
      <div style="font-size:18px;margin-top:8px;">${orderId}</div>
    </div>
    <div style="background:#fff;padding:15px;border-radius:8px;margin-bottom:12px;">
      <h3>📋 Informations du Client</h3>
      <p><strong>Nom:</strong> ${data.customerName || 'Non spécifié'}</p>
      <p><strong>Téléphone:</strong> ${data.customerPhone || 'Non spécifié'}</p>
      <p><strong>Adresse:</strong> ${data.customerAddress || 'Non spécifié'} — ${data.customerCommune || ''} — ${data.wilaya || ''}</p>
    </div>
    <div style="background:#fff;padding:15px;border-radius:8px;margin-bottom:12px;">
      <h3>🚚 Livraison</h3>
      <p><strong>Type:</strong> ${deliveryType}</p>
      <p><strong>Frais:</strong> ${formatNumber(data.deliveryCost || 0)} DA</p>
      <p><strong>Livraison gratuite:</strong> ${freeDeliveryText}</p>
    </div>
    <div style="background:#fff;padding:15px;border-radius:8px;margin-bottom:12px;">
      <h3>🛒 Détails des Articles</h3>
      <table style="width:100%;border-collapse:collapse;font-size:14px;">
        <thead>
          <tr>
            <th style="background:#4A6CF7;color:#fff;padding:10px;text-align:left;">Produit</th>
            <th style="background:#4A6CF7;color:#fff;padding:10px;">Quantité</th>
            <th style="background:#4A6CF7;color:#fff;padding:10px;text-align:right;">Prix Unitaire</th>
            <th style="background:#4A6CF7;color:#fff;padding:10px;text-align:right;">Total</th>
            <th style="background:#4A6CF7;color:#fff;padding:10px;">Taille</th>
            <th style="background:#4A6CF7;color:#fff;padding:10px;">Couleur</th>
          </tr>
        </thead>
        <tbody>${itemsHtml}</tbody>
      </table>
    </div>
    <div style="background:linear-gradient(135deg,#fff9e6,#ffe6cc);padding:12px;border-radius:8px;border:2px solid #FF9E16;">
      <p><strong>Sous-total:</strong> ${formatNumber(data.subtotal || 0)} DA</p>
      ${data.discount > 0 ? `<p><strong>Remise:</strong> -${formatNumber(data.discount || 0)} DA</p>` : ''}
      <p><strong>Frais livraison:</strong> ${formatNumber(data.deliveryCost || 0)} DA</p>
      <h3 style="color:#06D6A0;">TOTAL: ${formatNumber(data.total || 0)} DA</h3>
    </div>
    <p style="color:#666;font-size:12px;margin-top:18px;">Email généré automatiquement - ${new Date().toLocaleString('fr-FR')}</p>
  </body>
  </html>`;
}

// ---------- خريطة الألوان ----------
function getColorHex(colorName) {
  const colorMap = {
    "كما في الصورة": "#4A6CF7", "أبيض": "#ffffff", "أسود": "#000000", "رمادي": "#808080",
    "أزرق": "#3498db", "أحمر": "#e74c3c", "أخضر": "#2ecc71", "زهري": "#e84393",
    "بنفسجي": "#9b59b6", "Rouge": "#e74c3c", "Bleu": "#3498db", "Vert": "#2ecc71",
    "Rose": "#e84393", "Violet": "#9b59b6", "Blanc": "#ffffff", "Noir": "#000000", "Gris": "#808080"
  };
  return colorMap[colorName] || "#bdc3c7";
}
function formatNumber(number) {
  try { return new Intl.NumberFormat('fr-FR').format(Number(number || 0)); } catch (e) { return number; }
}

// ---------- اختبار محلي ----------
function testOrderSubmission() {
  const testData = {
    storeKey: 'test.pages.dev',
    storeName: 'TEST STORE',
    customerName: "Jean Dupont", customerPhone: "0550123456", customerAddress: "123 Rue de la République",
    customerCommune: "Alger Centre", wilaya: "Alger", deliveryType: "home", deliveryCost: 600,
    subtotal: 15000, discount: 1500, total: 14100, isFreeDelivery: false,
    items: [
      { name: "Chemise élégante", quantity: 2, price: 5000, finalPrice: 4500, size: "M", color: "Bleu" },
      { name: "Pantalon classique", quantity: 1, price: 6000, finalPrice: 6000, size: "42", color: "Noir" }
    ]
  };
  if (!PropertiesService.getScriptProperties().getProperty(ORDERS_API_SECRET_PROPERTY)) setupOrdersApi();
  return processOrderDirect(testData);
}

// ---------- GET ----------
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'active', message: 'Service Google Apps Script opérationnel!',
    timestamp: new Date().toLocaleString('fr-FR'), quota: MailApp.getRemainingDailyQuota()
  })).setMimeType(ContentService.MimeType.JSON);
}
function sendResponse(status, message, extra) {
  var payload = Object.assign({ status: status, message: message }, extra || {});
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// StoreMaster — Secure Multi-Tenant Orders API
// =====================================================================
function normalizeStoreKey(value) {
  return String(value || '').trim().toLowerCase()
    .replace(/^https?:\/\//, '').replace(/\/.*$/, '')
    .replace(/^\/+|\/+$/g, '');
}
function isValidOrdersApiSecret(secret) {
  const configured = String(PropertiesService.getScriptProperties().getProperty(ORDERS_API_SECRET_PROPERTY) || '').trim();
  return Boolean(configured && String(secret || '').trim() === configured);
}

const SHEET_HEADERS = [
  'StoreKey','Statut','RequestID','Date et Heure','ID Commande','Nom Client','Téléphone',
  'Commune','Adresse','Wilaya','Type Livraison','Articles Commandés','Quantité Totale',
  'Sous-total','Remise','Frais Livraison','Total Général','OrderData'
];
const SHEET_WIDTH = SHEET_HEADERS.length;
function getOrdersSheetWidth() { return SHEET_WIDTH; }

function getHeaderIndexes(sheet) {
  const headerRow = HEADER_ROW;
  const lastColumn = Math.max(sheet.getLastColumn(), SHEET_WIDTH);
  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0].map(v => String(v || '').trim());
  const index = {};
  headers.forEach((h, i) => { if (h) index[h] = i; });
  SHEET_HEADERS.forEach(h => { if (!(h in index)) index[h] = -1; });
  return index;
}

function initializeOrdersSheet(sheet) {
  try {
    sheet.clear();
    try { sheet.getRange(TITLE_ROW, 1, 1, sheet.getMaxColumns()).breakApart(); } catch (_) {}
    sheet.getRange(TITLE_ROW, 1, 1, SHEET_WIDTH).merge();
    sheet.getRange(TITLE_ROW, 1).setValue('🛍️ SHOPLIVE — TABLEAU DES COMMANDES')
      .setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(TITLE_ROW, 48);
    sheet.getRange(HEADER_ROW, 1, 1, SHEET_WIDTH).setValues([SHEET_HEADERS]);
    formatHeader(sheet);
    setColumnWidths(sheet);
    sheet.setFrozenRows(HEADER_ROW);
  } catch (err) { console.error('Erreur initializeOrdersSheet:', err); throw err; }
}

function ensureOrderSheetSchema(sheet) {
  if (sheet.getLastRow() === 0) { initializeOrdersSheet(sheet); return; }

  const maxCols = Math.max(sheet.getLastColumn(), SHEET_WIDTH);
  const row2 = sheet.getRange(HEADER_ROW, 1, 1, maxCols).getValues()[0].map(v => String(v || '').trim());
  const row1 = sheet.getRange(TITLE_ROW, 1, 1, maxCols).getValues()[0].map(v => String(v || '').trim());

  // 1) Existing original structure with headers on ROW 1 (as in the current sheet).
  // We migrate it safely to the multi-tenant structure without losing existing orders.
  if (row1[0] === 'Statut' && row1[1] === 'RequestID') {
    const oldHeaders = row1.slice(0, sheet.getLastColumn());
    const oldLastRow = sheet.getLastRow();
    const oldData = oldLastRow >= 2 ? sheet.getRange(2, 1, oldLastRow - 1, oldHeaders.length).getValues() : [];
    const oldIndex = {};
    oldHeaders.forEach((h, i) => { if (h) oldIndex[h] = i; });
    const val = (r, h) => oldIndex[h] === undefined ? '' : r[oldIndex[h]];
    const migrated = oldData.map(r => {
      const out = new Array(SHEET_WIDTH).fill('');
      // Existing historical rows do not have a reliable store identity.
      // Keep StoreKey blank rather than assigning them to the wrong store.
      out[0] = '';
      out[1] = val(r, 'Statut') || 'En attente';
      out[2] = val(r, 'RequestID');
      out[3] = val(r, 'Date et Heure');
      out[4] = val(r, 'ID Commande');
      out[5] = val(r, 'Nom Client');
      out[6] = val(r, 'Téléphone');
      out[7] = val(r, 'Commune');
      out[8] = val(r, 'Adresse');
      out[9] = val(r, 'Wilaya');
      out[10] = val(r, 'Type Livraison');
      out[11] = val(r, 'Articles Commandés');
      out[12] = val(r, 'Quantité Totale');
      out[13] = val(r, 'Sous-total');
      out[14] = val(r, 'Remise');
      out[15] = val(r, 'Frais Livraison');
      out[16] = val(r, 'Total Général');
      out[17] = '';
      return out;
    }).filter(r => String(r[2] || '').trim() !== '');
    sheet.clear();
    initializeOrdersSheet(sheet);
    if (migrated.length) sheet.getRange(DATA_START_ROW, 1, migrated.length, SHEET_WIDTH).setValues(migrated);
    for (let i = 0; i < migrated.length; i++) formatOrderRow(sheet, DATA_START_ROW + i);
    return;
  }
  // 2) Original SHOPLIVE structure with title row 1 + headers row 2, 16 columns.
  if (row2[0] === 'Statut' && row2[1] === 'RequestID') {
    sheet.insertColumnBefore(1);
    sheet.getRange(HEADER_ROW, 1).setValue('StoreKey');
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(HEADER_ROW, SHEET_WIDTH).setValue('OrderData');
  }
  // 3) Previous Multi-Tenant structure: headers on row 1.
  else if (row1[0] === 'StoreKey' && row1[1] === 'RequestID') {
    const oldHeaders = row1.slice(0, sheet.getLastColumn());
    const oldLastRow = sheet.getLastRow();
    const oldData = oldLastRow >= 2 ? sheet.getRange(2, 1, oldLastRow - 1, oldHeaders.length).getValues() : [];
    const oldIndex = {};
    oldHeaders.forEach((h, i) => { if (h) oldIndex[h] = i; });
    const val = (r, h) => oldIndex[h] === undefined ? '' : r[oldIndex[h]];
    const migrated = oldData.map(r => {
      const out = new Array(SHEET_WIDTH).fill('');
      out[0] = val(r, 'StoreKey');
      out[1] = val(r, 'Statut') || 'En attente';
      out[2] = val(r, 'RequestID');
      out[3] = val(r, 'Date et Heure');
      out[4] = val(r, 'ID Commande');
      out[5] = val(r, 'Nom Client');
      out[6] = val(r, 'Téléphone');
      out[7] = val(r, 'Commune');
      out[8] = val(r, 'Adresse');
      out[9] = val(r, 'Wilaya');
      out[10] = val(r, 'Type Livraison');
      out[11] = val(r, 'Articles Commandés');
      out[12] = val(r, 'Quantité Totale');
      out[13] = val(r, 'Sous-total');
      out[14] = val(r, 'Remise');
      out[15] = val(r, 'Frais Livraison');
      out[16] = val(r, 'Total Général');
      out[17] = val(r, 'OrderData');
      return out;
    }).filter(r => String(r[2] || '').trim() !== '');
    sheet.clear();
    initializeOrdersSheet(sheet);
    if (migrated.length) sheet.getRange(DATA_START_ROW, 1, migrated.length, SHEET_WIDTH).setValues(migrated);
    for (let i = 0; i < migrated.length; i++) formatOrderRow(sheet, DATA_START_ROW + i);
    return;
  }
  // 3) Unknown but empty-ish sheet.
  else if (row2[0] !== 'StoreKey' || row2[1] !== 'Statut') {
    if (sheet.getLastRow() <= 1) { initializeOrdersSheet(sheet); return; }
    throw new Error('Structure de la feuille Commandes non reconnue');
  }

  // Ensure the hidden technical OrderData column exists.
  const headersNow = sheet.getRange(HEADER_ROW, 1, 1, Math.max(sheet.getLastColumn(), SHEET_WIDTH)).getValues()[0].map(v => String(v || '').trim());
  if (headersNow.indexOf('OrderData') === -1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
  }
  // Normalize expected headers/positions.
  SHEET_HEADERS.forEach((h, i) => sheet.getRange(HEADER_ROW, i + 1).setValue(h));

  // Rebuild the title merge so adding StoreKey/OrderData never leaves a broken title range.
  try { sheet.getRange(TITLE_ROW, 1, 1, sheet.getMaxColumns()).breakApart(); } catch (_) {}
  try { sheet.getRange(TITLE_ROW, 1, 1, SHEET_WIDTH).merge(); } catch (_) {}
  sheet.getRange(TITLE_ROW, 1).setValue('🛍️ SHOPLIVE — TABLEAU DES COMMANDES')
    .setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E78')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(TITLE_ROW, 48);
  formatHeader(sheet);
  setColumnWidths(sheet);
  sheet.setFrozenRows(HEADER_ROW);
  addStatusValidationToSheet(sheet);
}
function addStatusValidationToSheet(sheet) {
  const idx = getHeaderIndexes(sheet);
  if (idx.Statut < 0) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(ORDER_STATUSES, true).setAllowInvalid(false).build();
  sheet.getRange(DATA_START_ROW, idx.Statut + 1, lastRow - DATA_START_ROW + 1, 1).setDataValidation(rule);
}
function addStatusDropdown(sheet, row) {
  const idx = getHeaderIndexes(sheet);
  const cell = sheet.getRange(row, idx.Statut + 1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(ORDER_STATUSES, true).setAllowInvalid(false).build();
  cell.setDataValidation(rule);
  if (!cell.getValue()) cell.setValue('En attente');
  applyStatusColor(sheet, row);
}
function applyStatusColor(sheet, row) {
  const idx = getHeaderIndexes(sheet);
  const cell = sheet.getRange(row, idx.Statut + 1);
  const status = String(cell.getValue() || '').trim();
  cell.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  if (status === 'Confirmée') cell.setBackground('#D9EAD3').setFontColor('#274E13');
  else if (status === 'En cours de livraison') cell.setBackground('#CFE2F3').setFontColor('#073763');
  else if (status === 'Livrée') cell.setBackground('#B6D7A8').setFontColor('#1B4332');
  else if (status === 'Retour') cell.setBackground('#F4CCCC').setFontColor('#990000');
  else cell.setBackground('#FFF2CC').setFontColor('#7F6000');
}
function formatHeader(sheet) {
  sheet.getRange(HEADER_ROW, 1, 1, SHEET_WIDTH).setBackground('#2F75B5').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true).setBorder(true,true,true,true,true,true,'#FFD966',SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(HEADER_ROW, 38);
}
function setColumnWidths(sheet) {
  const widths = {1:150,2:150,3:145,4:145,5:155,6:150,7:125,8:125,9:190,10:110,11:140,12:390,13:105,14:120,15:110,16:125,17:130,18:40};
  Object.keys(widths).forEach(c => sheet.setColumnWidth(Number(c), widths[c]));
}
function formatOrderRow(sheet, row) {
  try {
    const idx = getHeaderIndexes(sheet);
    sheet.getRange(row,1,1,SHEET_WIDTH).setVerticalAlignment('middle').setWrap(true)
      .setBorder(true,true,true,true,true,true,'#FFD966',SpreadsheetApp.BorderStyle.SOLID);
    addStatusDropdown(sheet,row);
    sheet.getRange(row,idx['Téléphone']+1).setNumberFormat('@').setHorizontalAlignment('center');
    sheet.getRange(row,idx['Date et Heure']+1).setNumberFormat('dd/MM/yyyy HH:mm').setHorizontalAlignment('center');
    sheet.getRange(row,idx['Quantité Totale']+1).setNumberFormat('0').setHorizontalAlignment('center');
    [idx['Sous-total']+1,idx.Remise+1,idx['Frais Livraison']+1,idx['Total Général']+1].forEach(c => sheet.getRange(row,c).setNumberFormat('#,##0 "DA"').setHorizontalAlignment('right'));
    sheet.getRange(row,idx['ID Commande']+1).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(row,idx['Articles Commandés']+1).setWrap(true).setVerticalAlignment('middle');
    sheet.getRange(row,idx.OrderData+1).setFontColor('#FFFFFF').setBackground('#FFFFFF');
    sheet.setRowHeight(row,100);
  } catch(err) { console.error('Erreur formatOrderRow:',err); }
}

function getOrdersSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  ensureOrderSheetSchema(sheet);
  return sheet;
}

function processOrderDirect(data) {
  const storeKey = normalizeStoreKey(data.storeKey);
  if (!storeKey) throw new Error('StoreKey obligatoire');
  const requestId = generateRequestId(data, storeKey);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (isRequestInSheet(requestId, storeKey)) return { success: true, message: 'Commande déjà enregistrée', requestId };
    const orderId = generateOrderId();
    saveToSpreadsheet(data, orderId, requestId, storeKey);

    if (!isEmailSentForRequest(requestId)) {
      const emailResult = sendSingleEmail(data, orderId, STORE_EMAIL, data.storeName || STORE_NAME);
      if (emailResult.success) markEmailAsSent(requestId);
    }
    markRequestAsCompleted(requestId);
    return { success: true, orderId, requestId, storeKey };
  } finally { try { lock.releaseLock(); } catch (_) {} }
}
function apiListOrders(body) {
  const storeKey = normalizeStoreKey(body.storeKey);
  if (!storeKey) throw new Error('StoreKey obligatoire');
  const sheet = getOrdersSheet();
  const idx = getHeaderIndexes(sheet);
  const lastRow = sheet.getLastRow();
  const orders = [];
  if (lastRow >= DATA_START_ROW) {
    const values = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, sheet.getLastColumn()).getValues();
    values.forEach((row, n) => {
      if (normalizeStoreKey(row[idx.StoreKey]) !== storeKey) return;
      const order = rowToOrder(row, idx, n + 2);
      if (matchesOrderFilters(order, body.filters || {})) orders.push(order);
    });
  }
  orders.sort((a,b) => (parseOrderDate(b.timestamp)?.getTime() || 0) - (parseOrderDate(a.timestamp)?.getTime() || 0));
  return { success: true, orders: orders, stats: buildOrderStats(orders) };
}
function apiGetOrder(body) {
  const storeKey = normalizeStoreKey(body.storeKey), orderId = String(body.orderId || '').trim();
  if (!storeKey || !orderId) throw new Error('storeKey et orderId obligatoires');
  const sheet = getOrdersSheet(), idx = getHeaderIndexes(sheet), lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { success:false, error:'Commande introuvable' };
  const values = sheet.getRange(DATA_START_ROW,1,lastRow-DATA_START_ROW+1,sheet.getLastColumn()).getValues();
  for (let i=0;i<values.length;i++) {
    if (normalizeStoreKey(values[i][idx.StoreKey]) === storeKey && String(values[i][idx['ID Commande']] || '') === orderId) return { success:true, order:rowToOrder(values[i],idx,i+2) };
  }
  return { success:false, error:'Commande introuvable' };
}
function apiUpdateOrderStatus(body) {
  const storeKey = normalizeStoreKey(body.storeKey), orderId = String(body.orderId || '').trim(), status = String(body.status || '').trim();
  if (!storeKey || !orderId || !status) throw new Error('Paramètres de mise à jour incomplets');
  if (ORDER_STATUSES.indexOf(status) === -1) throw new Error('Statut non autorisé');
  const sheet = getOrdersSheet(), idx = getHeaderIndexes(sheet), lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { success:false, error:'Commande introuvable' };
  const values = sheet.getRange(DATA_START_ROW,1,lastRow-DATA_START_ROW+1,sheet.getLastColumn()).getValues();
  for (let i=0;i<values.length;i++) {
    if (normalizeStoreKey(values[i][idx.StoreKey]) === storeKey && String(values[i][idx['ID Commande']] || '') === orderId) {
      const rowNumber = i + DATA_START_ROW;
      sheet.getRange(rowNumber, idx.Statut + 1).setValue(status);
      if (status === 'En attente') sheet.getRange(rowNumber, idx.Statut + 1).setBackground('#FFF2CC');
      else if (status === 'Confirmée') sheet.getRange(rowNumber, idx.Statut + 1).setBackground('#D9EAF7');
      else if (status === 'En cours de livraison') sheet.getRange(rowNumber, idx.Statut + 1).setBackground('#E8DDF7');
      else if (status === 'Livrée') sheet.getRange(rowNumber, idx.Statut + 1).setBackground('#D9EAD3');
      else if (status === 'Retour') sheet.getRange(rowNumber, idx.Statut + 1).setBackground('#F4CCCC');
      return { success:true, order:rowToOrder(sheet.getRange(rowNumber,1,1,sheet.getLastColumn()).getValues()[0],idx,rowNumber) };
    }
  }
  return { success:false, error:'Commande introuvable' };
}
function rowToOrder(row, idx, rowNumber) {
  const articlesText = String(row[idx['Articles Commandés']] || '');
  let items = [];
  if (idx.OrderData >= 0) {
    try {
      const parsed = JSON.parse(String(row[idx.OrderData] || '[]'));
      if (Array.isArray(parsed)) items = parsed;
    } catch (_) {}
  }
  return {
    row: rowNumber,
    storeKey: String(row[idx.StoreKey] || ''),
    requestId: String(row[idx.RequestID] || ''),
    timestamp: row[idx['Date et Heure']] instanceof Date ? Utilities.formatDate(row[idx['Date et Heure']], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(row[idx['Date et Heure']] || ''),
    orderId: String(row[idx['ID Commande']] || ''),
    customerName: String(row[idx['Nom Client']] || ''),
    customerPhone: String(row[idx['Téléphone']] || ''),
    customerCommune: String(row[idx.Commune] || ''),
    customerAddress: String(row[idx.Adresse] || ''),
    wilaya: String(row[idx.Wilaya] || ''),
    deliveryType: String(row[idx['Type Livraison']] || ''),
    itemsText: articlesText,
    items: items,
    totalPieces: Number(row[idx['Quantité Totale']] || 0),
    subtotal: Number(row[idx['Sous-total']] || 0),
    discount: Number(row[idx.Remise] || 0),
    deliveryCost: Number(row[idx['Frais Livraison']] || 0),
    total: Number(row[idx['Total Général']] || 0),
    status: String(row[idx.Statut] || 'En attente')
  };
}
function matchesOrderFilters(o, f) {
  const q = String(f.search || '').trim().toLowerCase();
  if (q && ![o.orderId,o.customerName,o.customerPhone,o.wilaya,o.customerCommune,o.itemsText].some(v => String(v || '').toLowerCase().includes(q))) return false;
  if (f.status && o.status !== f.status) return false;
  if (f.deliveryType) {
    const wanted = String(f.deliveryType).toLowerCase();
    if (wanted === 'desk' && !o.deliveryType.toLowerCase().includes('bureau')) return false;
    if (wanted === 'home' && !o.deliveryType.toLowerCase().includes('domicile')) return false;
  }
  const orderDate = parseOrderDate(o.timestamp);
  if (orderDate && f.dateFrom) {
    const from = new Date(String(f.dateFrom) + 'T00:00:00');
    if (orderDate < from) return false;
  }
  if (orderDate && f.dateTo) {
    const to = new Date(String(f.dateTo) + 'T23:59:59');
    if (orderDate > to) return false;
  }
  return true;
}
function parseOrderDate(value) {
  if (value instanceof Date) return value;
  const s = String(value || '').trim();
  let m = s.match(/^(\d{1,2})[\/](\d{1,2})[\/](\d{4})(?:,?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function buildOrderStats(orders) {
  return { total:orders.length, pending:orders.filter(o=>o.status==='En attente').length, confirmed:orders.filter(o=>o.status==='Confirmée').length, delivery:orders.filter(o=>o.status==='En cours de livraison').length, delivered:orders.filter(o=>o.status==='Livrée').length, returned:orders.filter(o=>o.status==='Retour').length };
}
function ordersApi(body) {
  if (!isValidOrdersApiSecret(body.secret)) return {success:false,error:'Accès non autorisé'};
  const action=String(body.action||'').trim();
  if (action==='createOrder') {
    const data=body.data||{}; data.storeKey=normalizeStoreKey(body.storeKey); data.storeName=String(body.storeName||'').trim();
    return processOrderDirect(data);
  }
  if (action==='list') return apiListOrders(body);
  if (action==='get' || action==='getOrder') return apiGetOrder(body);
  if (action==='updateStatus' || action==='updateOrderStatus') return apiUpdateOrderStatus(body);
  if (action==='setup') return setupOrdersApi();
  return {success:false,error:'Action inconnue'};
}
function setupOrdersApi() {
  const sheet=getOrdersSheet();
  let secret=PropertiesService.getScriptProperties().getProperty(ORDERS_API_SECRET_PROPERTY);
  if (!secret) { secret=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,''); PropertiesService.getScriptProperties().setProperty(ORDERS_API_SECRET_PROPERTY,secret); }
  return {success:true,message:'Sheet prête et Secret configuré',secret:secret,columns:sheet.getLastColumn()};
}

// ---------- دوال إدارة عامة ----------
function clearAllProcessedRequests() {
  try { const props=PropertiesService.getScriptProperties(); const keys=Object.keys(props.getProperties()||{}); keys.forEach(k=>props.deleteProperty(k)); console.log('🧹 Toutes les requêtes ont été nettoyées'); return {success:true}; }
  catch(err){console.error('Erreur nettoyage:',err);return {success:false,error:err.toString()};}
}
function showProcessedRequests() {
  try { const props=PropertiesService.getScriptProperties(); const p=props.getProperties(); console.log('📋 Requêtes actuelles:',p); return p; }
  catch(err){console.error('Erreur affichage:',err);return {};}
}
