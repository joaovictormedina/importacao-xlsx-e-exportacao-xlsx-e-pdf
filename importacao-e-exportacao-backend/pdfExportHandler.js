const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const exportPdf = async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'imported_file.xlsx'));

        const worksheet = workbook.getWorksheet('Cotacao');
        if (!worksheet) {
            return res.status(400).json({ error: 'A aba "Cotacao" não foi encontrada.' });
        }

        // Salva o arquivo Excel com a aba "Cotacao"
        const outputPath = path.join(__dirname, 'exported_cotacao.xlsx');
        await workbook.xlsx.writeFile(outputPath);

        // Envia o arquivo Excel para o cliente
        res.sendFile(outputPath);
    } catch (error) {
        console.error('Erro ao exportar o arquivo em PDF:', error);
        res.status(500).json({ error: 'Erro ao exportar o arquivo em PDF.' });
    }
};

module.exports = { exportPdf };
