// const express = require('express');
// const xlsx = require('xlsx');
// const path = require('path');
// const fs = require('fs');

// const router = express.Router();

// router.post('/', (req, res) => {
//     const { file1, file2 } = req.files;

//     if (!file1 || !file2) {
//         return res.status(400).send('Veuillez télécharger deux fichiers Excel.');
//     }

//     try {
//         // Lire les fichiers Excel
//         const workbook1 = xlsx.read(file1.data, { type: 'buffer' });
//         const workbook2 = xlsx.read(file2.data, { type: 'buffer' });

//         const sheetName1 = workbook1.SheetNames[0];
//         const sheetName2 = workbook2.SheetNames[0];

//         const worksheet1 = workbook1.Sheets[sheetName1];
//         const worksheet2 = workbook2.Sheets[sheetName2];

//         // Convertir en JSON
//         const data1 = xlsx.utils.sheet_to_json(worksheet1);
//         const data2 = xlsx.utils.sheet_to_json(worksheet2);

//         // Fusionner les données
//         const mergedData = [...data1, ...data2];

//         // Créer un nouveau workbook et une feuille
//         const newWorkbook = xlsx.utils.book_new();
//         const newWorksheet = xlsx.utils.json_to_sheet(mergedData);
        
//         // Ajouter la feuille au nouveau workbook
//         xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Merged Data');

//         // Définir le chemin de sortie
//         const outputDir = path.join(__dirname);
//         const outputFilePath = path.join(outputDir, 'Merged_Report.xlsx');

//         // Écrire le nouveau fichier Excel
//         xlsx.writeFile(newWorkbook, outputFilePath);

//         // Envoyer le fichier fusionné au client
//         res.download(outputFilePath, 'Merged_Report.xlsx', (err) => {
//             if (err) {
//                 console.error('Erreur lors de l\'envoi du fichier:', err);
//                 res.status(500).send('Erreur lors de l\'envoi du fichier.');
//             }

//             // Optionnel : supprimer le fichier après l'envoi
//             fs.unlink(outputFilePath, (unlinkErr) => {
//                 if (unlinkErr) {
//                     console.error('Erreur lors de la suppression du fichier temporaire:', unlinkErr);
//                 }
//             });
//         });
//     } catch (error) {
//         console.error('Erreur lors de la fusion des fichiers Excel:', error);
//         res.status(500).send('Erreur lors de la fusion des fichiers Excel.');
//     }
// });

// module.exports = router;
