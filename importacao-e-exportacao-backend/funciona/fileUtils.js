const XLSX = require('xlsx');
const fs = require('fs');

// Função para salvar o arquivo importado
const saveImportedFile = (workbook, outputPath) => {
    console.log(`Salvando arquivo importado em: ${outputPath}`); // Adicione este log
    XLSX.writeFile(workbook, outputPath);
};

// Função para salvar o arquivo exportado
const saveExportedFile = (workbook, outputPath) => {
    console.log(`Salvando arquivo exportado em: ${outputPath}`); // Adicione este log
    XLSX.writeFile(workbook, outputPath);
};

module.exports = { saveImportedFile, saveExportedFile };
