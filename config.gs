// config.gs - Configuration globale et gestion des onglets Google Sheet

// Variable globale pour le spreadsheet
let spreadsheet = null;

// Initialiser la feuille de calcul et créer les onglets
function initializeSpreadsheet() {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Liste des onglets à créer
    const sheetsToCreate = [
        'Config',
        'Admins',
        'Professeurs',
        'Étudiants',
        'Contact',
        'Inscription_En_Attente'
    ];

    // Créer les onglets s'ils n'existent pas
    sheetsToCreate.forEach(sheetName => {
        if (!getSheetByName(sheetName)) {
            spreadsheet.insertSheet(sheetName);
        }
    });

    // Initialiser les en-têtes seulement
    initializeHeaders();
}

// Initialiser les en-têtes des onglets
function initializeHeaders() {
    initializeConfigHeaders();
    initializeAdminsHeaders();
    initializeProfesseursHeaders();
    initializeEtudiantsHeaders();
    initializeContactHeaders();
    initializeInscriptionHeaders();
}

// Initialiser les en-têtes de Config
function initializeConfigHeaders() {
    const sheet = getSheetByName('Config');
    if (!sheet) return;

    // Vérifier si les en-têtes existent
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Clé', 'Valeur']);
        // Formater les en-têtes
        const headerRange = sheet.getRange(1, 1, 1, 2);
        headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
    }
}

// Initialiser les en-têtes d'Admins
function initializeAdminsHeaders() {
    const sheet = getSheetByName('Admins');
    if (!sheet) return;

    // Vérifier si les en-têtes existent
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Nom', 'Email', 'Mot de passe', 'Rôle', 'Date d\'ajout']);
        // Formater les en-têtes
        const headerRange = sheet.getRange(1, 1, 1, 5);
        headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
    }
}

// Initialiser les en-têtes de Professeurs
function initializeProfesseursHeaders() {
    const sheet = getSheetByName('Professeurs');
    if (!sheet) return;

    // Vérifier si les en-têtes existent
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Nom', 'Prénom', 'Email', 'Mot de passe', 'Classe', 'Matière', 'Téléphone', 'Date d\'ajout']);
        // Formater les en-têtes
        const headerRange = sheet.getRange(1, 1, 1, 8);
        headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
    }
}

// Initialiser les en-têtes d'Étudiants
function initializeEtudiantsHeaders() {
    const sheet = getSheetByName('Étudiants');
    if (!sheet) return;

    // Vérifier si les en-têtes existent
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Matricule', 'Nom', 'Prénom', 'Email', 'Classe', 'Téléphone', 'Statut', 'Date d\'inscription']);
        // Formater les en-têtes
        const headerRange = sheet.getRange(1, 1, 1, 8);
        headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
    }
}

// Initialiser les en-têtes de Contact
function initializeContactHeaders() {
    const sheet = getSheetByName('Contact');
    if (!sheet) return;

    // Vérifier si les en-têtes existent
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Clé', 'Valeur']);
        // Formater les en-têtes
        const headerRange = sheet.getRange(1, 1, 1, 2);
        headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
    }
}

// Initialiser les en-têtes d'Inscription En Attente
function initializeInscriptionHeaders() {
    const sheet = getSheetByName('Inscription_En_Attente');
    if (!sheet) return;

    // Vérifier si les en-têtes existent
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Nom', 'Prénom', 'Email', 'Classe', 'Téléphone', 'Matricule Temporaire', 'Date d\'inscription', 'Statut']);
        // Formater les en-têtes
        const headerRange = sheet.getRange(1, 1, 1, 8);
        headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
    }
}

// Obtenir une feuille par son nom
function getSheetByName(name) {
    try {
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        return spreadsheet.getSheetByName(name);
    } catch (e) {
        return null;
    }
}

// Obtenir toutes les données du spreadsheet
function getAllData() {
    if (!spreadsheet) {
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }

    return {
        config: getConfigData(),
        admins: getAdminsData(),
        profs: getProfesseursData(),
        students: getEtudiantsData(),
        pending: getPendingRegistrations()
    };
}

// Obtenir les données de configuration
function getConfigData() {
    const sheet = getSheetByName('Config');
    if (!sheet || sheet.getLastRow() < 2) return {};

    const data = sheet.getDataRange().getValues();
    const config = {};

    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
            config[data[i][0]] = data[i][1];
        }
    }

    return config;
}

// Obtenir les données des admins
function getAdminsData() {
    const sheet = getSheetByName('Admins');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const admins = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][1]) {  // Vérifier que Nom et Email existent
            admins.push({
                nom: data[i][0],
                email: data[i][1],
                password: data[i][2],
                role: data[i][3] || 'Admin'
            });
        }
    }

    return admins;
}

// Obtenir les données des professeurs
function getProfesseursData() {
    const sheet = getSheetByName('Professeurs');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const profs = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][2]) {  // Vérifier que Nom et Email existent
            profs.push({
                nom: data[i][0],
                prenom: data[i][1] || '',
                email: data[i][2],
                password: data[i][3],
                classe: data[i][4] || '',
                matiere: data[i][5] || '',
                telephone: data[i][6] || ''
            });
        }
    }

    return profs;
}

// Obtenir les données des étudiants
function getEtudiantsData() {
    const sheet = getSheetByName('Étudiants');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const students = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][1]) {  // Vérifier que Matricule et Nom existent
            students.push({
                matricule: data[i][0],
                nom: data[i][1],
                prenom: data[i][2] || '',
                email: data[i][3] || '',
                classe: data[i][4] || '',
                telephone: data[i][5] || '',
                statut: data[i][6] || 'Actif'
            });
        }
    }

    return students;
}

// Obtenir les inscriptions en attente
function getPendingRegistrations() {
    const sheet = getSheetByName('Inscription_En_Attente');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const pending = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][1]) {  // Vérifier que Nom et Prénom existent
            pending.push({
                nom: data[i][0],
                prenom: data[i][1],
                email: data[i][2] || '',
                classe: data[i][3] || '',
                telephone: data[i][4] || '',
                matricule: data[i][5] || '',
                date_inscription: data[i][6] || '',
                statut: data[i][7] || 'En attente'
            });
        }
    }

    return pending;
}

// Initialiser automatiquement au déploiement
function onOpen(e) {
    initializeSpreadsheet();
}
