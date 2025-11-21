/**
 * CONFIGURATION GLOBALE
 * Centralise les noms des feuilles et les index de colonnes.
 */
const CONFIG = {
  // Configuration pour les Laptops
  LAPTOP: {
    type: 'LAPTOP',
    sheetName: "Stock laptop",
    historySheet: "Historique laptop",
    prefix: "PT-",
    // Mapping des colonnes (Nom logique : Nom exact de l'en-tête dans le Sheet)
    cols: {
      id: "PT-code",
      tag: "Service Tag",
      user: "USER",
      status: "STATUS",
      location: "LOCATION",
      company: "COMPANY",
      model: "MODEL"
    }
  },
  
  // Configuration pour les Monitors
  MONITOR: {
    type: 'MONITOR',
    sheetName: "Stock monitor",
    historySheet: "Historique monitor",
    prefix: "MON-",
    cols: {
      id: "MON-code",
      tag: "Service Tag",
      user: "USER",
      status: "STATUS",
      location: "LOCATION",
      company: "COMPANY",
      model: "MODEL MONITOR" // Attention au nom spécifique ici
    }
  },

  // Feuille contenant les listes déroulantes
  DROPDOWN_SHEET: "Kind of"
};

// Helper pour récupérer la bonne config
function getConfigByType(type) {
  return type === 'MONITOR' ? CONFIG.MONITOR : CONFIG.LAPTOP;
}