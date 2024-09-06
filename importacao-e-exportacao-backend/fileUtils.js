const XLSX = require('xlsx');

// Função para salvar o arquivo importado
const saveImportedFile = (workbook, outputPath) => {
    XLSX.writeFile(workbook, outputPath);
};

module.exports = { saveImportedFile };
