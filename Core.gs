/**
 * Gère les opérations sur les feuilles de calcul de manière robuste.
 */
class SheetManager {
  constructor(sheetName) {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.ss.getSheetByName(sheetName);
    if (!this.sheet) throw new Error(`Feuille introuvable : ${sheetName}`);
    
    this.lastRow = this.sheet.getLastRow();
    this.lastCol = this.sheet.getLastColumn();
    
    // Chargement des en-têtes pour le mapping dynamique
    if (this.lastRow > 0) {
      this.headers = this.sheet.getRange(1, 1, 1, this.lastCol).getValues()[0];
      this.headerMap = this.createHeaderMap(this.headers);
    } else {
      this.headers = [];
      this.headerMap = {};
    }
  }

  createHeaderMap(headers) {
    let map = {};
    headers.forEach((h, i) => map[String(h).trim()] = i);
    return map;
  }

  getColumnIndex(colName) {
    if (this.headerMap[colName] === undefined) return -1;
    return this.headerMap[colName];
  }

  /**
   * Transforme un objet JS {status: "Active"} en ligne Sheet [..., "Active", ...]
   */
  objectToRow(dataObj) {
    let row = new Array(this.headers.length).fill("");
    
    for (let key in dataObj) {
      // Tentative de mapping automatique (clé camelCase -> En-tête UPPERCASE)
      let headerName = key.toUpperCase(); 
      
      // Mappings manuels pour les exceptions
      if(key === 'serviceTag') headerName = 'Service Tag';
      if(key === 'purchaseDate') headerName = 'PURCHASE DATE';
      if(key === 'decomissionDate') headerName = 'DECOMISSION DATE';
      if(key === 'poNumber') headerName = 'PO NUMBER';
      if(key === 'hardDisk') headerName = 'HARD DISK';
      if(key === 'ticketNumber') continue; // On ne stocke pas le ticket dans le stock

      let colIndex = this.getColumnIndex(headerName);
      
      // Si pas trouvé, on cherche la clé exacte
      if (colIndex === -1) colIndex = this.getColumnIndex(key);

      if (colIndex !== -1) {
        row[colIndex] = dataObj[key];
      }
    }
    return row;
  }

  appendData(dataObj) {
    const row = this.objectToRow(dataObj);
    this.sheet.appendRow(row);
    return row;
  }
  
  // Génère un ID unique (PT-XXXXX) avec verrouillage pour éviter les doublons
  generateNextId(idColName, prefix) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000); // Attend max 10sec
      
      const colIndex = this.getColumnIndex(idColName);
      if (colIndex === -1) throw new Error("Colonne ID introuvable (" + idColName + ")");

      const data = this.sheet.getRange(2, colIndex + 1, this.sheet.getLastRow(), 1).getValues();
      let maxId = 0;
      
      data.flat().forEach(val => {
        if (val && String(val).startsWith(prefix)) {
          let num = parseInt(String(val).replace(prefix, ""), 10);
          if (!isNaN(num) && num > maxId) maxId = num;
        }
      });
      
      return prefix + String(maxId + 1).padStart(5, '0');
      
    } catch (e) {
      throw e;
    } finally {
      lock.releaseLock();
    }
  }
  
  findRowIndexByColumn(colName, value) {
    const colIndex = this.getColumnIndex(colName);
    if (colIndex === -1) return -1;
    
    const data = this.sheet.getDataRange().getValues();
    // i starts at 1 to skip headers, return i+1 for 1-based index
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIndex]).toLowerCase().trim() === String(value).toLowerCase().trim()) return i + 1; 
    }
    return -1;
  }

  updateCell(rowIndex, colName, value) {
    const colIndex = this.getColumnIndex(colName);
    if (colIndex !== -1) {
      this.sheet.getRange(rowIndex, colIndex + 1).setValue(value);
    }
  }
}

/**
 * Helper pour écrire dans l'historique
 */
function logToHistory(config, action, serviceTag, details) {
  const db = new SheetManager(config.historySheet);
  const userEmail = Session.getActiveUser().getEmail();
  
  let historyData = {
    'Who do it': userEmail,
    'what he do': action,
    'wich device': serviceTag, // Sic: orthographe originale conservée
    'when': new Date(),
    'Ticket number': details.ticket || '',
    'old users': details.oldUser || '',
    'new users': details.newUser || '',
    'old status': details.oldStatus || '',
    'new status': details.newStatus || '',
    'old location': details.oldLocation || '',
    'new location': details.newLocation || '',
    'company': details.company || ''
  };
  
  db.appendData(historyData);
}