function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Stock Tools V2')
    .addItem('Ajouter Laptops', 'showAddDevicesForm')
    .addItem('Ajouter Ecrans', 'showAddMonitorsForm')
    .addSeparator()
    .addItem('Mettre à jour Laptop', 'showUpdateDeviceForm')
    .addItem('Mettre à jour Ecran', 'showUpdateMonitorForm')
    .addSeparator()
    .addItem('Remplacer Laptop', 'showReplaceDeviceForm')
    .addItem('Remplacer Ecran', 'showReplaceMonitorForm')
    .addToUi();
}

function showModal(fileName, title, width, height, params) {
  var template = HtmlService.createTemplateFromFile(fileName);
  if (params) {
    Object.keys(params).forEach(key => template[key] = params[key]);
  }
  var html = template.evaluate().setWidth(width).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

// --- ADD ---
function showAddDevicesForm() {
  var params = getDropdownOptions(); 
  params.assetType = 'LAPTOP'; 
  showModal('Add_Asset_Form', 'Ajouter Laptops', 1100, 800, params);
}

function showAddMonitorsForm() {
  var params = getDropdownOptions();
  params.assetType = 'MONITOR';
  showModal('Add_Asset_Form', 'Ajouter Moniteurs', 1100, 800, params);
}

// --- UPDATE ---
function showUpdateDeviceForm() {
  var params = getDropdownOptions();
  params.assetType = 'LAPTOP';
  showModal('Update_Asset_Form', 'Mettre à jour Laptop', 600, 750, params);
}

function showUpdateMonitorForm() {
  var params = getDropdownOptions();
  params.assetType = 'MONITOR';
  showModal('Update_Asset_Form', 'Mettre à jour Monitor', 600, 750, params);
}

// --- REPLACE ---
function showReplaceDeviceForm() {
  showModal('Replace_Asset_Form', 'Remplacer Laptop', 600, 600, { assetType: 'LAPTOP' });
}

function showReplaceMonitorForm() {
  showModal('Replace_Asset_Form', 'Remplacer Monitor', 600, 600, { assetType: 'MONITOR' });
}