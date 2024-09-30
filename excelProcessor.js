const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Fonction pour obtenir le numéro de la semaine actuelle
function getCurrentWeek() {
    const today = new Date();
    const start = new Date(today.getFullYear(), 0, 1);
    const diff = (today - start) + ((start.getTimezoneOffset() - today.getTimezoneOffset()) * 60 * 1000);
    const oneWeek = 1000 * 60 * 60 * 24 * 7;
    const week = Math.floor(diff / oneWeek) + 1;
    return week;
}

// Fonction pour ajouter du style à une cellule (bordure, fond, alignement)
function addCellStyle(cell) {
    cell.border = {
        top: { style: 'medium' },
        bottom: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' }
    };
    cell.alignment = {
        horizontal: 'center',
        vertical: 'center'
    };
}

exports.processExcelFile = async (excelBuffer, res) => {
    try {
        console.log('Début du traitement du fichier Excel.');

        // Lire le fichier Excel depuis un buffer
        // const workbook = new ExcelJS.Workbook();
        // await workbook.xlsx.load(excelBuffer);
        // console.log('Fichier Excel lu avec succès.');

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Résultats');

        // Vérifier si le workbook contient des feuilles
        // if (!workbook.worksheets || workbook.worksheets.length === 0) {
        //     throw new Error("Le fichier Excel ne contient aucune feuille.");
        // }

        // const sheet = workbook.getWorksheet(1); // Première feuille
        // if (!sheet) {
        //     throw new Error("Impossible d'accéder à la première feuille du fichier Excel.");
        // }
        // console.log('Nom de la feuille :', sheet.name);

        // Initialiser les compteurs
        let countStartsWithI = 0;
        let countStartsWithS = 0;
        let priorityCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        let countNetwork = 0;
        let countSystem = 0;
        let countSupervisionNagios = 0;

        // Parcourir les lignes et effectuer les comptages nécessaires
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Ignorer l'en-tête

            const rfcNumber = row.getCell('A').text;
            const priorityValue = row.getCell('B').value;
            const groupValue = row.getCell('C').text;
            const categorie1Value = row.getCell('D').text;

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

        console.log('Comptage terminé avec succès.');

        // Obtenir la semaine actuelle
        const currentWeek = getCurrentWeek();
        const currentWeekColumn = `Semaine ${currentWeek}`;

        // Créer un nouveau workbook et une nouvelle feuille pour les résultats
        const newWorkbook = new ExcelJS.Workbook();
        const resultSheet = newWorkbook.addWorksheet('Résultats');

        // Ajouter les données avec les résultats des compteurs
        const resultData = [
            { label: "Incidents", value: countStartsWithI },
            { label: "Demandes", value: countStartsWithS },
            { label: "P1", value: priorityCounts[1] },
            { label: "P2", value: priorityCounts[2] },
            { label: "P3", value: priorityCounts[3] },
            { label: "P4", value: priorityCounts[4] },
            { label: "P5", value: priorityCounts[5] },
            { label: "Network", value: countNetwork },
            { label: "System", value: countSystem },
            { label: "Supervision Nagios", value: countSupervisionNagios }
        ];

        resultSheet.columns = [
            { header: 'Indicateur / Semaine', key: 'label', width: 30 },
            { header: currentWeekColumn, key: 'value', width: 15 }
        ];

        resultSheet.addRows(resultData);

        // Appliquer des styles aux en-têtes et aux cellules
        resultSheet.eachRow((row, rowNumber) => {
            row.eachCell((cell) => {
                if (rowNumber === 1) {
                    // Style pour l'en-tête
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'D3D3D3' }
                    };
                    cell.font = { bold: true };
                }

                // Appliquer des bordures et un centrage pour toutes les cellules
                addCellStyle(cell);
            });
        });

        // Enregistrer le fichier Excel avec les résultats
        const outputFilePath = path.join(__dirname, 'resultat.xlsx');
        await newWorkbook.xlsx.writeFile(outputFilePath);

        console.log('Fichier Excel avec les résultats généré avec succès.');

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
};
