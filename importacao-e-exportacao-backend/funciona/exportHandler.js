// exportHandler.js
const XLSX = require('xlsx');
const path = require('path');
const { saveExportedFile } = require('./fileUtils');

// Função para exportar os dados
const exportFile = (req, res) => {
    try {
        const data = req.body.data;

        if (!data) {
            return res.status(400).json({ error: 'Dados não fornecidos para exportação.' });
        }

        // Cria uma nova planilha
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

        // Salva o arquivo exportado
        const outputPath = path.join(__dirname, 'exported_file.xlsx');
        saveExportedFile(wb, outputPath);

        // Envia o arquivo para o cliente
        res.sendFile(outputPath);
    } catch (error) {
        console.error('Erro ao exportar o arquivo:', error);
        res.status(500).json({ error: 'Erro ao exportar o arquivo.' });
    }
};

module.exports = { exportFile };
