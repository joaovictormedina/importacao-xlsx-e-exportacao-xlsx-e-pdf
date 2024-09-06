// exportHandler.js
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Função para exportar o arquivo
const exportFile = async (req, res) => {
    try {
        const inputPath = path.join(__dirname, 'imported_file.xlsx');

        // Verifica se o arquivo existe
        if (!fs.existsSync(inputPath)) {
            return res.status(404).json({ error: 'Arquivo importado não encontrado.' });
        }

        // Lê o arquivo
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(inputPath);

        // Salva o arquivo exportado
        const outputPath = path.join(__dirname, 'exported_file.xlsx');
        await workbook.xlsx.writeFile(outputPath);

        // Envia o arquivo para o cliente
        res.sendFile(outputPath);
    } catch (error) {
        console.error('Erro ao exportar o arquivo:', error);
        res.status(500).json({ error: 'Erro ao exportar o arquivo.' });
    }
};

module.exports = { exportFile };
