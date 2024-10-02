const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const mergeExcelFiles = (file1, file2, outputFilePath) => {
    return new Promise((resolve, reject) => {
        try {
            // Lire les deux fichiers Excel
            const workbook1 = xlsx.read(file1.data, { type: 'buffer' });
            const workbook2 = xlsx.read(file2.data, { type: 'buffer' });

            // Obtenir les noms de feuilles des deux fichiers
            const sheetName1 = workbook1.SheetNames[0];
            const sheetName2 = workbook2.SheetNames[0];
            
            // Extraire les données sous forme de JSON
            const data1 = xlsx.utils.sheet_to_json(workbook1.Sheets[sheetName1]);
            const data2 = xlsx.utils.sheet_to_json(workbook2.Sheets[sheetName2]);

            // Obtenir le nom de la colonne B du fichier 2
            const columnsF2 = Object.keys(data2[0]);
            const columnBName = columnsF2[1]; // Supposons que la deuxième colonne est B

            // Fusionner les données des deux fichiers
            const mergedData = [];

            // Ajouter les données du premier fichier
            data1.forEach((row, index) => {
                // Ajouter une nouvelle ligne avec les données du fichier 1
                const newRow = { ...row };

                // Si une donnée existe pour la colonne B du fichier 2, l'ajouter
                if (data2[index] && data2[index][columnBName]) { 
                    newRow[columnBName] = data2[index][columnBName];
                } else {
                    // Si pas de donnée disponible dans le fichier 2, ajouter une valeur vide
                    newRow[columnBName] = '';
                }

                mergedData.push(newRow);
            });

            // Créer un nouveau classeur et ajouter les données fusionnées
            const newWorkbook = xlsx.utils.book_new();
            const newWorksheet = xlsx.utils.json_to_sheet(mergedData);
            xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Fusionné');

            // Enregistrer le fichier fusionné
            xlsx.writeFile(newWorkbook, outputFilePath);
            resolve(outputFilePath);
        } catch (error) {
            reject(error);
        }
    });
};

module.exports = { mergeExcelFiles };
