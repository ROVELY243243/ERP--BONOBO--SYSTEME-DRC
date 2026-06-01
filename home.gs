// home.gs - Backend pour la page d'accueil

// Fonction principale de déploiement
function doGet() {
    return HtmlService.createHtmlOutputFromFile('index')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Obtenir la configuration de l'école
function getSchoolConfig() {
    const config = getConfigData();
    
    return {
        nom_ecole: config['nom_ecole'] || 'ERP Bonobo Système DRC',
        description: config['description'] || 'Plateforme de gestion scolaire moderne',
        adresse: config['adresse'] || 'Kinshasa, RDC',
        telephone: config['telephone'] || '+243 XXX XXX XXX',
        email: config['email'] || 'contact@bonobo-erp.com',
        site_web: config['site_web'] || 'www.bonobo-erp.com',
        couleur_primaire: config['couleur_primaire'] || '#0066CC',
        logo_url: config['logo_url'] || 'https://via.placeholder.com/150'
    };
}

// Login Personnel (Admin ou Professeur)
function loginPersonnel(email, password) {
    // Vérifier d'abord les admins
    const admins = getAdminsData();
    const admin = admins.find(a => a.email === email && a.password === password);
    
    if (admin) {
        return {
            success: true,
            user: admin,
            role: 'admin'
        };
    }

    // Vérifier ensuite les professeurs
    const profs = getProfesseursData();
    const prof = profs.find(p => p.email === email && p.password === password);
    
    if (prof) {
        return {
            success: true,
            user: prof,
            role: 'professeur'
        };
    }

    return {
        success: false,
        message: 'Email ou mot de passe incorrect'
    };
}

// Trouver un étudiant par matricule
function findStudent(matricule) {
    const students = getEtudiantsData();
    const student = students.find(s => s.matricule === matricule);
    
    if (student) {
        return {
            success: true,
            student: student
        };
    }

    return {
        success: false,
        message: 'Étudiant non trouvé'
    };
}

// Enregistrer un nouvel étudiant
function registerStudent(studentData) {
    // Vérifier que l'email n'existe pas déjà
    const students = getEtudiantsData();
    const existingStudent = students.find(s => s.email === studentData.email);
    
    if (existingStudent) {
        return {
            success: false,
            message: 'Cet email est déjà enregistré'
        };
    }

    // Générer un matricule temporaire
    const matricule = generateMatricule();
    
    // Ajouter à l'onglet Inscription_En_Attente
    const sheet = getSheetByName('Inscription_En_Attente');
    if (!sheet) {
        return {
            success: false,
            message: 'Erreur: feuille non trouvée'
        };
    }

    const date = new Date();
    sheet.appendRow([
        studentData.nom,
        studentData.prenom,
        studentData.email,
        studentData.classe,
        studentData.telephone,
        matricule,
        date,
        'En attente'
    ]);

    return {
        success: true,
        matricule: matricule,
        message: 'Inscription enregistrée. En attente de confirmation par l\'administrateur.'
    };
}

// Générer un matricule unique
function generateMatricule() {
    const year = new Date().getFullYear();
    const students = getEtudiantsData();
    const pending = getPendingRegistrations();
    
    // Compter les étudiants de l'année en cours
    const allStudents = [...students, ...pending];
    const thisYearStudents = allStudents.filter(s => {
        if (s.matricule) {
            return s.matricule.includes('-' + year + '-');
        }
        return false;
    });
    
    const count = thisYearStudents.length + 1;
    const countStr = String(count).padStart(5, '0');
    
    return 'MAT-' + year + '-' + countStr;
}

// Obtenir une feuille par son nom
function getSheetByName(name) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        return ss.getSheetByName(name);
    } catch (e) {
        return null;
    }
}

// Importer les fonctions de config.gs
// Ces fonctions sont définies dans config.gs et sont accessibles ici
