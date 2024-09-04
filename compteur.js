const xlsx = require('xlsx');
const process = require('process');

// Récupérer les noms de fichiers passés en argument
const inputFileName = process.argv[2] || 'details.xlsx';
const outputFileName = process.argv[3] || 'resultat.xlsx';

// Lire le fichier Excel source
const workbook = xlsx.readFile(inputFileName);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convertir la feuille en format JSON pour faciliter le traitement
const data = xlsx.utils.sheet_to_json(worksheet);

// Initialiser les compteurs
let countStartsWithI = 0;
let countStartsWithS = 0;
let priorityCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
let countNetwork = 0;
let countSystem = 0;
let countSupervisionNagios = 0;

// Parcourir les lignes et effectuer les comptages nécessaires
data.forEach(row => {
    const rfcNumber = row['RFC_NUMBER'];
    const priorityValue = row['PRIORITY_VALUE'];
    const groupValue = row['Groupe'];
    const categorie1Value = row['Catégorie 1'];

    if (typeof rfcNumber === 'string') {
        if (rfcNumber.startsWith('I')) {
            countStartsWithI++;
        } else if (rfcNumber.startsWith('S')) {
            countStartsWithS++;
        }
    }

    const priority = Number(priorityValue);
    if (!isNaN(priority) && priorityCounts.hasOwnProperty(priority)) {
        priorityCounts[priority]++;
    }

    if (typeof groupValue === 'string') {
        const lowerGroupValue = groupValue.toLowerCase();
        if (lowerGroupValue.includes('network')) {
            countNetwork++;
        } else if (lowerGroupValue.includes('system')) {
            countSystem++;
        }
    }

    if (typeof categorie1Value === 'string' && categorie1Value === 'Supervision Nagios') {
        countSupervisionNagios++;
    }
});

// Créer une nouvelle feuille avec les résultats
const resultData = [
    { "Lettre": "I", "Nombre d'éléments": countStartsWithI },
    { "Lettre": "S", "Nombre d'éléments": countStartsWithS },
    { "Priorité": 1, "Nombre d'occurrences": priorityCounts[1] },
    { "Priorité": 2, "Nombre d'occurrences": priorityCounts[2] },
    { "Priorité": 3, "Nombre d'occurrences": priorityCounts[3] },
    { "Priorité": 4, "Nombre d'occurrences": priorityCounts[4] },
    { "Priorité": 5, "Nombre d'occurrences": priorityCounts[5] },
    { "Groupe": "Network", "Nombre d'occurrences": countNetwork },
    { "Groupe": "System", "Nombre d'occurrences": countSystem },
    { "Catégorie 1": "Supervision Nagios", "Nombre d'occurrences": countSupervisionNagios }
];

// Convertir les résultats en une feuille de calcul
const resultSheet = xlsx.utils.json_to_sheet(resultData);

// Créer un nouveau classeur Excel et ajouter la feuille de résultats
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, resultSheet, 'Résultats');

// Enregistrer le classeur Excel avec les résultats dans un nouveau fichier
xlsx.writeFile(newWorkbook, outputFileName);

console.log(`Les résultats ont été enregistrés dans le fichier: ${outputFileName}`);
