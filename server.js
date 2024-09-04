const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Middleware pour gérer les fichiers uploadés
app.use(fileUpload());

// Servir la page HTML pour l'upload
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Route pour gérer le téléchargement du fichier Excel
app.post('/upload', (req, res) => {
    if (!req.files || !req.files.excelFile) {
        return res.status(400).send('Aucun fichier n\'a été téléchargé.');
    }

    // Obtenir le fichier téléchargé en tant que buffer
    const excelFile = req.files.excelFile;

    try {
        // Lire le fichier Excel en tant que buffer
        const workbook = xlsx.read(excelFile.data, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        // Afficher les données lues pour vérification
        console.log('Données lues du fichier Excel:', data);

        // Initialiser les compteurs
        let countStartsWithI = 0;
        let countStartsWithS = 0;
        let priorityCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        let countNetwork = 0;
        let countSystem = 0;
        let countSupervisionNagios = 0;

        // Parcourir les lignes et effectuer les comptages nécessaires
        data.forEach(row => {
            console.log('Ligne traitée:', row); // Afficher chaque ligne traitée

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
            { "Élément": "Incidents", "Total": countStartsWithI },
            { "Élément": "Demandes", "Total": countStartsWithS },
            { "Élément": "P1", "Total": priorityCounts[1] },
            { "Élément": "P2", "Total": priorityCounts[2] },
            { "Élément": "P3", "Total": priorityCounts[3] },
            { "Élément": "P4", "Total": priorityCounts[4] },
            { "Élément": "P5", "Total": priorityCounts[5] },
            { "Élément": "Network", "Total": countNetwork },
            { "Élément": "System", "Total": countSystem },
            { "Élément": "Supervision Nagios", "Total": countSupervisionNagios }
        ];

        const resultSheet = xlsx.utils.json_to_sheet(resultData, { header: ["Élément", "Total"] });
        const newWorkbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(newWorkbook, resultSheet, 'Résultats');

        // Enregistrer le classeur Excel avec les résultats dans un fichier temporaire
        const outputFilePath = path.join(__dirname, 'resultat.xlsx');
        xlsx.writeFile(newWorkbook, outputFilePath);

        // Envoyer le fichier généré au client
        res.download(outputFilePath, 'resultat.xlsx', (err) => {
            if (err) {
                console.error('Erreur lors de l\'envoi du fichier:', err);
                res.status(500).send('Erreur lors du traitement du fichier.');
            }

            // Supprimer le fichier temporaire après envoi
            fs.unlink(outputFilePath, (unlinkErr) => {
                if (unlinkErr) {
                    console.error('Erreur lors de la suppression du fichier temporaire:', unlinkErr);
                }
            });
        });
    } catch (error) {
        console.error('Erreur lors du traitement du fichier Excel:', error);
        res.status(500).send('Erreur lors du traitement du fichier Excel.');
    }
});

// Démarrer le serveur
app.listen(PORT, () => {
    console.log(`Serveur démarré sur http://localhost:${PORT}`);
});
