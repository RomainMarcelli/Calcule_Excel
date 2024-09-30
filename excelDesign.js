const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Fonction pour obtenir la semaine actuelle (norme ISO 8601)
function getISOWeekNumber() {
    const date = new Date();
    const target = new Date(date.valueOf());
    const dayNumber = (date.getUTCDay() + 6) % 7;

    // Lundi de la semaine actuelle
    target.setUTCDate(target.getUTCDate() - dayNumber + 3);
    const firstThursday = new Date(target.getFullYear(), 0, 4);

    // Calcul du numéro de semaine
    const diff = target - firstThursday;
    const oneWeek = 1000 * 60 * 60 * 24 * 7;
    return 1 + Math.round(diff / oneWeek);
}

// Fonction pour ajouter des bordures et du centrage aux cellules avec des styles optionnels
const addCellStyle = (cell, customStyles = {}) => {
    // Fusionner les styles par défaut avec les styles spécifiques à la cellule
    const alignment = customStyles.alignment || { vertical: 'middle', horizontal: 'center' };
    const border = customStyles.border || {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'medium' }
    };

    cell.alignment = alignment;
    cell.border = border;
};

async function generateStyledExcel(countStartsWithI, countStartsWithS, priorityCounts, countNetwork, countSystem, countSupervisionNagios, outputFilePath) {
    // Créer le dossier s'il n'existe pas
    const dir = path.dirname(outputFilePath);
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }

    // Créer une nouvelle instance du classeur Excel
    const newWorkbook = new ExcelJS.Workbook();
    const resultSheet = newWorkbook.addWorksheet('Résultats Stylisés');

    // Obtenir la semaine actuelle (ISO 8601)
    const currentWeek = getISOWeekNumber();

    // Créer les colonnes pour l'Excel
    const currentWeekColumn = `Semaine ${currentWeek}`;
    resultSheet.columns = [
        { header: 'Indicateur / Semaine', key: 'label', width: 50 },
        { header: currentWeekColumn, key: 'value', width: 15 }
    ];

    // Ajouter la ligne avec "Nombre Total tickets ouverts (lundi au dimanche)"
    const totalRow = resultSheet.addRow({
        label: 'Nombre Total tickets ouverts\n (lundi au dimanche)',
        value: countStartsWithI + countStartsWithS
    });

    // Activer le retour automatique à la ligne pour la cellule A2
    totalRow.getCell(1).alignment = { wrapText: true };

    // Appliquer un style à la cellule A2
    const cellA2 = resultSheet.getCell('A2');
    cellA2.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D3D3D3' } // Gris clair
    };
    cellA2.font = { bold: true };
    cellA2.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
    addCellStyle(cellA2, {
        border: {
            top: { style: 'medium' },   // Bordure en haut de type 'medium'
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'medium' }
        }
    });

    // Créer les données des résultats
    const resultData = [
        { label: "Incidents", value: countStartsWithI },
        { label: "Demandes", value: countStartsWithS },
        { label: 'Priorité P1 etc…\n (Incidents & Demandes)', value: '' },
        { label: "P1", value: priorityCounts[1] },
        { label: "P2", value: priorityCounts[2] },
        { label: "P3", value: priorityCounts[3] },
        { label: "P4", value: priorityCounts[4] },
        { label: "P5", value: priorityCounts[5] },
        { label: "Détails par groupes (uniquement incidents)", value: '' },
        { label: "Network", value: countNetwork },
        { label: "System", value: countSystem },
        { label: "Supervision ouvert (lundi au dimanche)", value: countSupervisionNagios },
        { label: "Prise en charge supervision", value: '' },
        { label: "SLA: pris en charge en moins de 30 minutes", value: '' },
        { label: "Détails de résolution - Incidents", value: '' },
        { label: "P2: Temps de traitement < 2h", value: '' },
        { label: "P3: Temps de traitement < 8h", value: '' },
        { label: "P4: Temps de traitement < 3j", value: '' },
        { label: "P5: Temps de traitement < 5j", value: '' },
        { label: "Escalade Incidents", value: '' },
        { label: "Nombre Total Incidents escaladé", value: '' },
        { label: "-N3 Système", value: '' }
    ];

    // Ajouter les lignes de données
    resultSheet.addRows(resultData);

    // Appliquer les styles aux cellules spécifiques
    const styleSpecificCells = [
        { cell: 'A5', color: 'D3D3D3', border: { top: 'medium', left: 'thick', bottom: 'thick', right: 'medium' } },
        { cell: 'A11', color: 'D3D3D3' },
        { cell: 'A15', color: 'D3D3D3' },
        { cell: 'A16', color: 'DAF2D0' },
        { cell: 'A17', color: 'D3D3D3' },
        { cell: 'A18', color: 'DAF2D0' },
        { cell: 'A19', color: 'DAF2D0' },
        { cell: 'A20', color: 'DAF2D0' },
        { cell: 'A21', color: 'DAF2D0' },
        { cell: 'A22', color: 'D3D3D3' }
    ];

    styleSpecificCells.forEach(({ cell, color, border }) => {
        const currentCell = resultSheet.getCell(cell);
        currentCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color }
        };
        currentCell.font = { bold: false };
        currentCell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        addCellStyle(currentCell, { border });
    });

    // Appliquer des styles aux en-têtes et aux cellules restantes
    resultSheet.eachRow((row, rowNumber) => {
        row.eachCell((cell) => {
            if (rowNumber === 1) {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'D3D3D3' }
                };
                cell.font = { bold: true };
            }
            addCellStyle(cell);
        });
    });

    const cellA5 = resultSheet.getCell('A5');
    // Ajouter le texte en gras
    cellA5.font = { bold: true };
    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA5, {
        border: {
            top: { style: 'medium' },   // Bordure en haut de type 'medium'
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'medium' }
        }
    });
    const cellA17 = resultSheet.getCell('A17');
    // Ajouter le texte en gras
    cellA17.font = { bold: true };
    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA17, {
        border: {
            top: { style: 'medium' },   // Bordure en haut de type 'medium'
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'medium' }
        }
    });

    const cellA11 = resultSheet.getCell('A11');
    // Ajouter le texte en gras
    cellA11.font = { bold: true };
    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA11, {
        border: {
            top: { style: 'medium' },   // Bordure en haut de type 'medium'
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'medium' }
        }
    });
    const cellA15 = resultSheet.getCell('A15');
    // Ajouter le texte en gras
    cellA15.font = { bold: true };
    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA15, {
        border: {
            top: { style: 'medium' },   // Bordure en haut de type 'medium'
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'medium' }
        }
    });
    const cellA22 = resultSheet.getCell('A22');
    // Ajouter le texte en gras
    cellA22.font = { bold: true };
    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA22, {
        border: {
            top: { style: 'medium' },   // Bordure en haut de type 'medium'
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'medium' }
        }
    });

    const cellA3 = resultSheet.getCell('A3');
    addCellStyle(cellA3, {
        alignment: { horizontal: 'left' },
    });
    const cellA4 = resultSheet.getCell('A4');
    addCellStyle(cellA4, {
        alignment: { horizontal: 'left' },
    });
    const cellA6 = resultSheet.getCell('A6');
    addCellStyle(cellA6, {
        alignment: { horizontal: 'left' },
    });
    const cellA7 = resultSheet.getCell('A7');
    addCellStyle(cellA7, {
        alignment: { horizontal: 'left' },
    });
    const cellA8 = resultSheet.getCell('A8');
    addCellStyle(cellA8, {
        alignment: { horizontal: 'left' },
    });
    const cellA9 = resultSheet.getCell('A9');
    addCellStyle(cellA9, {
        alignment: { horizontal: 'left' },
    });
    const cellA10 = resultSheet.getCell('A10');
    addCellStyle(cellA10, {
        alignment: { horizontal: 'left' },
    });
    const cellA12 = resultSheet.getCell('A12');
    addCellStyle(cellA12, {
        alignment: { horizontal: 'left' },
    });
    const cellA13 = resultSheet.getCell('A13');
    addCellStyle(cellA13, {
        alignment: { horizontal: 'left' },
    });
    const cellA14 = resultSheet.getCell('A14');
    addCellStyle(cellA14, {
        alignment: { horizontal: 'left' },
    });
    const cellA16 = resultSheet.getCell('A16');
    addCellStyle(cellA16, {
        alignment: { horizontal: 'left' },
    });
    const cellA23 = resultSheet.getCell('A23');
    addCellStyle(cellA23, {
        alignment: { horizontal: 'left' },
    });
    const cellA24 = resultSheet.getCell('A24');
    addCellStyle(cellA24, {
        alignment: { horizontal: 'left' },
    });

    const cellA18 = resultSheet.getCell('A18');
    cellA18.value = {
        richText: [
            { text: 'P2', font: { bold: true } },
            { text: ': Temps de traitement < 2h', font: { bold: false } }
        ]
    };
    cellA18.alignment = { vertical: 'middle', horizontal: 'left' };

    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA18, {
        alignment: { horizontal: 'left' },
    });

    // Ajouter la ligne pour P3
    const cellA19 = resultSheet.getCell('A19');
    cellA19.value = {
        richText: [
            { text: 'P3', font: { bold: true } },
            { text: '\n', font: { bold: false } }, // Saut de ligne
            { text: ': Temps de traitement < 8h', font: { bold: false } }
        ]
    };
    cellA19.alignment = { 
        vertical: 'middle', 
        horizontal: 'left', 
        wrapText: true // Autorise les retours à la ligne
    };

    resultSheet.getRow(19).height = 30;

    const cellA20 = resultSheet.getCell('A20');
    cellA20.value = {
        richText: [
            { text: 'P4', font: { bold: true } },
            { text: ': Temps de traitement < 3j', font: { bold: false } }
        ]
    };
    cellA20.alignment = { vertical: 'middle', horizontal: 'left' };

    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA20, {
        alignment: { horizontal: 'left' },
    });

    const cellA21 = resultSheet.getCell('A21');
    cellA21.value = {
        richText: [
            { text: 'P5', font: { bold: true } },
            { text: ': Temps de traitement < 5j', font: { bold: false } }
        ]
    };
    cellA21.alignment = { vertical: 'middle', horizontal: 'left' };

    // Ajouter un style avec bordure supérieure 'medium'
    addCellStyle(cellA21, {
        alignment: { horizontal: 'left' },
    });


    // Enregistrer le fichier Excel avec les résultats
    await newWorkbook.xlsx.writeFile(outputFilePath);

    console.log('Fichier Excel avec les résultats généré avec succès.');
}

module.exports = { generateStyledExcel };
