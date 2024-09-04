const express = require('express');
const fileUpload = require('express-fileupload');
const ExcelJS = require('exceljs');
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
app.post('/upload', async (req, res) => {
    if (!req.files || !req.files.excelFile) {
        return res.status(400).send('Aucun fichier n\'a été téléchargé.');
    }

    // Obtenir le fichier téléchargé en tant que buffer
    const excelFile = req.files.excelFile;

    try {
        // Lire le fichier Excel en tant que buffer
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(excelFile.data);
        const worksheet = workbook.worksheets[0];

        // Initialiser les compteurs
        let countStartsWithI = 0;
        let countStartsWithS = 0;
        let priorityCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        let countNetwork = 0;
        let countSystem = 0;
        let countSupervisionNagios = 0;

        // Parcourir les lignes et effectuer les comptages nécessaires
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            if (rowNumber === 1) return; // Ignorer l'en-tête

            const rfcNumber = row.getCell(1).value;
            const priorityValue = row.getCell(2).value;
            const groupValue = row.getCell(3).value;
            const categorie1Value = row.getCell(4).value;

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
        const newWorkbook = new ExcelJS.Workbook();
        const resultSheet = newWorkbook.addWorksheet('Résultats');

        // Ajouter des en-têtes de colonne
        resultSheet.addRow(['Lettre', 'Nombre d\'éléments']);
        resultSheet.addRow(['I', countStartsWithI]);
        resultSheet.addRow(['S', countStartsWithS]);

        resultSheet.addRow(['Priorité', 'Nombre d\'occurrences']);
        for (let i = 1; i <= 5; i++) {
            resultSheet.addRow([i, priorityCounts[i]]);
        }

        resultSheet.addRow(['Groupe', 'Nombre d\'éléments']);
        resultSheet.addRow(['Network', countNetwork]);
        resultSheet.addRow(['System', countSystem]);

        resultSheet.addRow(['Catégorie', 'Nombre d\'occurrences']);
        resultSheet.addRow(['Supervision Nagios', countSupervisionNagios]);

        // Enregistrer le classeur Excel avec les résultats dans un fichier temporaire
        const outputFilePath = path.join(__dirname, 'resultat.xlsx');
        await newWorkbook.xlsx.writeFile(outputFilePath);

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
