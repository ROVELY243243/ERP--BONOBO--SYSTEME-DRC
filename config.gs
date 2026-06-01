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

    // Initialiser le contenu des onglets
    initializeConfigSheet();
    initializeAdminsSheet();
    initializeProfesseursSheet();
    initializeEtudiantsSheet();
    initializeContactSheet();
    initializeInscriptionSheet();
}

// Obtenir une feuille par son nom
function getSheetByName(name) {
    try {
        return spreadsheet.getSheetByName(name);
    } catch (e) {
        return null;
    }
}

// Initialiser l'onglet Config
function initializeConfigSheet() {
    const sheet = getSheetByName('Config');
    if (!sheet) return;

    // Vérifier si déjà initialisé
    if (sheet.getLastRow() > 0) return;

    const headers = ['Clé', 'Valeur'];
    const data = [
        ['nom_ecole', 'ERP Bonobo Système DRC'],
        ['description', 'Plateforme de gestion scolaire moderne'],
        ['adresse', 'Kinshasa, République Démocratique du Congo'],
        ['telephone', '+243 XXX XXX XXX'],
        ['email', 'contact@bonobo-erp.com'],
        ['site_web', 'www.bonobo-erp.com'],
        ['couleur_primaire', '#0066CC'],
        ['clotaire_inscription', '31/12/2026'],
        ['annee_scolaire', '2025-2026'],
        ['logo_url', 'https://via.placeholder.com/150']
    ];

    sheet.appendRow(headers);
    data.forEach(row => sheet.appendRow(row));

    // Formater les en-têtes
    const headerRange = sheet.getRange(1, 1, 1, 2);
    headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
}

// Initialiser l'onglet Admins
function initializeAdminsSheet() {
    const sheet = getSheetByName('Admins');
    if (!sheet) return;

    // Vérifier si déjà initialisé
    if (sheet.getLastRow() > 0) return;

    const headers = ['Nom', 'Email', 'Mot de passe', 'Rôle', 'Date d\'ajout'];
    const data = [
        ['Admin Principal', 'admin@bonobo.com', 'admin123', 'Super Admin', new Date()]
    ];

    sheet.appendRow(headers);
    data.forEach(row => sheet.appendRow(row));

    // Formater les en-têtes
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
}

// Initialiser l'onglet Professeurs
function initializeProfesseursSheet() {
    const sheet = getSheetByName('Professeurs');
    if (!sheet) return;

    // Vérifier si déjà initialisé
    if (sheet.getLastRow() > 0) return;

    const headers = ['Nom', 'Prénom', 'Email', 'Mot de passe', 'Classe', 'Matière', 'Téléphone', 'Date d\'ajout'];
    const data = [
        ['Prof', 'Exemple', 'prof@bonobo.com', 'prof123', '6ème A', 'Mathématiques', '+243 XXX XXX XXX', new Date()]
    ];

    sheet.appendRow(headers);
    data.forEach(row => sheet.appendRow(row));

    // Formater les en-têtes
    const headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
}

// Initialiser l'onglet Étudiants
function initializeEtudiantsSheet() {
    const sheet = getSheetByName('Étudiants');
    if (!sheet) return;

    // Vérifier si déjà initialisé
    if (sheet.getLastRow() > 0) return;

    const headers = ['Matricule', 'Nom', 'Prénom', 'Email', 'Classe', 'Téléphone', 'Statut', 'Date d\'inscription'];
    const data = [
        ['MAT-2024-001', 'Étudiant', 'Exemple', 'etudiant@bonobo.com', '6ème A', '+243 XXX XXX XXX', 'Actif', new Date()]
    ];

    sheet.appendRow(headers);
    data.forEach(row => sheet.appendRow(row));

    // Formater les en-têtes
    const headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
}

// Initialiser l'onglet Contact
function initializeContactSheet() {
    const sheet = getSheetByName('Contact');
    if (!sheet) return;

    // Vérifier si déjà initialisé
    if (sheet.getLastRow() > 0) return;

    const headers = ['Clé', 'Valeur'];
    const data = [
        ['Adresse', 'Kinshasa, RDC'],
        ['Téléphone Principal', '+243 XXX XXX XXX'],
        ['Email Principal', 'contact@bonobo-erp.com'],
        ['Site Web', 'www.bonobo-erp.com'],
        ['Heures d\'ouverture', 'Lun-Ven: 8h-16h'],
        ['Responsable', 'Directeur de l\'école'],
        ['Téléphone Responsable', '+243 XXX XXX XXX']
    ];

    sheet.appendRow(headers);
    data.forEach(row => sheet.appendRow(row));

    // Formater les en-têtes
    const headerRange = sheet.getRange(1, 1, 1, 2);
    headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
}

// Initialiser l'onglet Inscription En Attente
function initializeInscriptionSheet() {
    const sheet = getSheetByName('Inscription_En_Attente');
    if (!sheet) return;

    // Vérifier si déjà initialisé
    if (sheet.getLastRow() > 0) return;

    const headers = ['Nom', 'Prénom', 'Email', 'Classe', 'Téléphone', 'Matricule Temporaire', 'Date d\'inscription', 'Statut'];
    
    sheet.appendRow(headers);

    // Formater les en-têtes
    const headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setBackground('#0066CC').setFontColor('white').setFontWeight('bold');
}

// Fonction pour obtenir toutes les données du spreadsheet
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
    if (!sheet) return {};

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
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const admins = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
            admins.push({
                nom: data[i][0],
                email: data[i][1],
                password: data[i][2],
                role: data[i][3]
            });
        }
    }

    return admins;
}

// Obtenir les données des professeurs
function getProfesseursData() {
    const sheet = getSheetByName('Professeurs');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const profs = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
            profs.push({
                nom: data[i][0],
                prenom: data[i][1],
                email: data[i][2],
                password: data[i][3],
                classe: data[i][4],
                matiere: data[i][5],
                telephone: data[i][6]
            });
        }
    }

    return profs;
}

// Obtenir les données des étudiants
function getEtudiantsData() {
    const sheet = getSheetByName('Étudiants');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const students = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
            students.push({
                matricule: data[i][0],
                nom: data[i][1],
                prenom: data[i][2],
                email: data[i][3],
                classe: data[i][4],
                telephone: data[i][5],
                statut: data[i][6]
            });
        }
    }

    return students;
}

// Obtenir les inscriptions en attente
function getPendingRegistrations() {
    const sheet = getSheetByName('Inscription_En_Attente');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const pending = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
            pending.push({
                nom: data[i][0],
                prenom: data[i][1],
                email: data[i][2],
                classe: data[i][3],
                telephone: data[i][4],
                matricule: data[i][5],
                date_inscription: data[i][6],
                statut: data[i][7]
            });
        }
    }

    return pending;
}

// Initialiser automatiquement au déploiement
function onOpen(e) {
    initializeSpreadsheet();
}
