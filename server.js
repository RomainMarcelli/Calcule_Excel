const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const { generateStyledExcel } = require('./excelDesign');
const {mergeExcelFiles} = require('./merge'); // Importer la fonction de design

const app = express();
const PORT = 3000;

// Middleware pour gérer les fichiers uploadés
app.use(fileUpload());

// Route pour afficher la page HTML de l'upload
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Route pour gérer l'upload et calculer les résultats
app.post('/upload', (req, res) => {
    if (!req.files || !req.files.excelFile) {
        return res.status(400).send('Aucun fichier téléchargé.');
    }

    const excelFile = req.files.excelFile;

    try {
        // Lire le fichier Excel
        const workbook = xlsx.read(excelFile.data, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        // Faire les calculs
        let countStartsWithI = 0;
        let countStartsWithS = 0;
        let priorityCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        let countNetwork = 0;
        let countSystem = 0;
        let countSupervisionNagios = 0;

        data.forEach(row => {
            const rfcNumber = row['RFC_NUMBER'];
            const priorityValue = row['PRIORITY_VALUE'];
            const groupValue = row['Groupe'];
            const categorie1Value = row['Catégorie 1'];

            if (typeof rfcNumber === 'string') {
                if (rfcNumber.startsWith('I')) countStartsWithI++;
                else if (rfcNumber.startsWith('S')) countStartsWithS++;
            }

            const priority = Number(priorityValue);
            if (!isNaN(priority) && priorityCounts.hasOwnProperty(priority)) {
                priorityCounts[priority]++;
            }

            if (typeof groupValue === 'string') {
                const lowerGroupValue = groupValue.toLowerCase();
                if (lowerGroupValue.includes('network')) countNetwork++;
                else if (lowerGroupValue.includes('system')) countSystem++;
            }

            if (typeof categorie1Value === 'string' && categorie1Value === 'Supervision Nagios') {
                countSupervisionNagios++;
            }
        });

        // Dossier où le fichier sera sauvegardé
        const outputDir = path.join(__dirname); // Change ce chemin si tu veux un autre dossier
        const outputFilePath = path.join(outputDir, 'Reporting_Nhood.xlsx');

        // Créer le dossier s'il n'existe pas
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        // Passer les résultats au fichier de design pour générer l'Excel stylisé
        generateStyledExcel(countStartsWithI, countStartsWithS, priorityCounts, countNetwork, countSystem, countSupervisionNagios, outputFilePath)
            .then(() => {
                // Envoyer le fichier Excel stylisé en réponse
                res.download(outputFilePath, 'Reporting_Nhood.xlsx', (err) => {
                    if (err) {
                        console.error('Erreur lors de l\'envoi du fichier:', err);
                        res.status(500).send('Erreur lors du traitement du fichier.');
                    }

                    // Supprimer le fichier temporaire après téléchargement
                    fs.unlink(outputFilePath, (unlinkErr) => {
                        if (unlinkErr) {
                            console.error('Erreur lors de la suppression du fichier temporaire:', unlinkErr);
                        }
                    });
                });
            })
            .catch((err) => {
                console.error('Erreur lors de la génération de l\'Excel:', err);
                res.status(500).send('Erreur lors de la génération du fichier Excel.');
            });
    } catch (error) {
        console.error('Erreur lors du traitement du fichier Excel:', error);
        res.status(500).send('Erreur lors du traitement du fichier Excel.');
    }
});

// Nouvelle route pour fusionner deux fichiers
app.post('/merge', (req, res) => {
    if (!req.files || !req.files.excelFile1 || !req.files.excelFile2) {
        return res.status(400).send('Deux fichiers sont nécessaires pour la fusion.');
    }

    const excelFile1 = req.files.excelFile1;
    const excelFile2 = req.files.excelFile2;
    const outputMergeFilePath = path.join(__dirname, 'Reporting_Merged.xlsx');

    mergeExcelFiles(excelFile1, excelFile2, outputMergeFilePath)
        .then(outputPath => {
            res.download(outputPath, 'Reporting_Merged.xlsx', (err) => {
                if (err) {
                    console.error('Erreur lors de l\'envoi du fichier fusionné:', err);
                    res.status(500).send('Erreur lors du traitement du fichier fusionné.');
                }

                // Supprimer le fichier temporaire après téléchargement
                fs.unlink(outputPath, (unlinkErr) => {
                    if (unlinkErr) {
                        console.error('Erreur lors de la suppression du fichier temporaire fusionné:', unlinkErr);
                    }
                });
            });
        })
        .catch(error => {
            console.error('Erreur lors du traitement des fichiers Excel:', error);
            res.status(500).send('Erreur lors du traitement des fichiers Excel.');
        });
});

// Démarrer le serveur
app.listen(PORT, () => {
    console.log(`Serveur démarré sur http://localhost:${PORT}`);
});
