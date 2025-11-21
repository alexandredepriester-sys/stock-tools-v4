/* === FONCTIONS PUBLIQUES (Appelées par le HTML) === */

// --- AJOUT ---
function addAsset(data, type) {
  const config = getConfigByType(type);
  const db = new SheetManager(config.sheetName);
  
  const newId = db.generateNextId(config.cols.id, config.prefix);
  data[config.cols.id] = newId; 
  
  db.appendData(data);
  
  logToHistory(config, 'Add new device', data.serviceTag, {
    newUser: data.user,
    newStatus: data.status,
    newLocation: data.location,
    company: data.company
  });
  
  return "Succès";
}

function addLaptops(dataArray) { dataArray.forEach(item => addAsset(item, 'LAPTOP')); }
function addMonitors(dataArray) { dataArray.forEach(item => addAsset(item, 'MONITOR')); }


// --- MISE A JOUR ---
function updateAssetData(formData, type) {
  const config = getConfigByType(type);
  const db = new SheetManager(config.sheetName);
  
  const rowIndex = db.findRowIndexByColumn(config.cols.tag, formData.serviceTag);
  if (rowIndex === -1) throw new Error("Service Tag introuvable : " + formData.serviceTag);
  
  // Mise à jour des champs clés
  if(formData.user !== undefined) db.updateCell(rowIndex, config.cols.user, formData.user);
  if(formData.status !== undefined) db.updateCell(rowIndex, config.cols.status, formData.status);
  if(formData.location !== undefined) db.updateCell(rowIndex, config.cols.location, formData.location);
  if(formData.company !== undefined) db.updateCell(rowIndex, config.cols.company, formData.company);
  if(formData.observation !== undefined) db.updateCell(rowIndex, "OBSERVATION", formData.observation);
  
  // TODO: Ajouter ici les autres champs si nécessaire (Purchase Date, etc.)

  logToHistory(config, 'Update', formData.serviceTag, {
    newUser: formData.user,
    newStatus: formData.status,
    ticket: formData.ticketNumber,
    newLocation: formData.location
  });
}

function updateDevice(data) { updateAssetData(data, 'LAPTOP'); }
function updateMonitor(data) { updateAssetData(data, 'MONITOR'); }


// --- REMPLACEMENT (SWAP) ---
function replaceAssetProcess(oldTag, newTag, observation, ticket, type) {
  const config = getConfigByType(type);
  const db = new SheetManager(config.sheetName);
  
  const oldRowIdx = db.findRowIndexByColumn(config.cols.tag, oldTag);
  const newRowIdx = db.findRowIndexByColumn(config.cols.tag, newTag);
  
  if (oldRowIdx === -1 || newRowIdx === -1) throw new Error("Service Tag introuvable (Vérifiez l'ancien et le nouveau).");
  
  // Récupération info Ancien
  const sheet = db.sheet;
  const userColIdx = db.getColumnIndex(config.cols.user) + 1;
  const companyColIdx = db.getColumnIndex(config.cols.company) + 1;
  
  const oldUser = sheet.getRange(oldRowIdx, userColIdx).getValue();
  const company = sheet.getRange(oldRowIdx, companyColIdx).getValue();
  
  // 1. UPDATE ANCIEN -> Stock
  db.updateCell(oldRowIdx, config.cols.user, "");
  db.updateCell(oldRowIdx, config.cols.status, "In stock");
  db.updateCell(oldRowIdx, "OBSERVATION", observation);
  
  // 2. UPDATE NOUVEAU -> User
  db.updateCell(newRowIdx, config.cols.user, oldUser);
  db.updateCell(newRowIdx, config.cols.status, "Assigned");
  
  // 3. HISTORIQUE
  logToHistory(config, 'Device Replacement (Old)', oldTag, {
    oldUser: oldUser,
    newStatus: "In stock",
    ticket: ticket,
    company: company
  });
  
  logToHistory(config, 'Device Acquisition (New)', newTag, {
    newUser: oldUser,
    newStatus: "Assigned",
    ticket: ticket,
    company: company
  });
  
  return "Remplacement effectué.";
}

function replaceDeviceInSheet(o, n, obs, tick) { return replaceAssetProcess(o, n, obs, tick, 'LAPTOP'); }
function replaceMonitorInSheet(o, n, obs, tick) { return replaceAssetProcess(o, n, obs, tick, 'MONITOR'); }


/* === HELPERS (Listes, Autocomplete) === */

function getDropdownOptions() {
  const db = new SheetManager(CONFIG.DROPDOWN_SHEET);
  const data = db.sheet.getDataRange().getValues();
  const headers = data.shift();
  let options = {};
  
  headers.forEach((h, i) => {
    options[h] = data.map(row => row[i]).filter(cell => cell !== "");
  });
  return options;
}

function getServiceTagsByType(type) {
  const config = getConfigByType(type);
  const db = new SheetManager(config.sheetName);
  const colIndex = db.getColumnIndex(config.cols.tag);
  const raw = db.sheet.getRange(2, colIndex + 1, db.lastRow, 1).getValues().flat();
  return raw.filter(String);
}

// Alias pour compatibilité HTML
function getServiceTags() { return getServiceTagsByType('LAPTOP'); }
function getServiceTagsMonitor() { return getServiceTagsByType('MONITOR'); }
function getExistingServiceTags() { return getServiceTagsByType('LAPTOP'); }
function getExistingServiceTagsMonitor() { return getServiceTagsByType('MONITOR'); }

// Alias pour compatibilité Users (si utilisé)
function getUsersAdd(query) {
  // On prend la liste de tous les users de la feuille Stock Laptop comme référence
  const db = new SheetManager(CONFIG.LAPTOP.sheetName);
  const colIndex = db.getColumnIndex(CONFIG.LAPTOP.cols.user);
  const users = db.sheet.getRange(2, colIndex + 1, db.lastRow, 1).getValues().flat();
  return users.filter(u => u && u.toLowerCase().includes(query.toLowerCase()));
}
// Récupération des données complètes pour Update Form
function getDeviceDataByServiceTag(tag, type) {
  const config = getConfigByType(type);
  const db = new SheetManager(config.sheetName);
  const idx = db.findRowIndexByColumn(config.cols.tag, tag);
  
  if (idx === -1) return null;
  
  // Conversion de la ligne en objet
  const rowData = db.sheet.getRange(idx, 1, 1, db.lastCol).getValues()[0];
  const headers = db.headers;
  let obj = {};
  
  headers.forEach((h, i) => {
    let val = rowData[i];
    // Formatage Date
    if (val instanceof Date) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    obj[h] = val;
  });
  
  return obj;
}