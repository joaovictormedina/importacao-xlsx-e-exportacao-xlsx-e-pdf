// importHandler.js
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');

// Configuração do multer para lidar com uploads de arquivos
const upload = multer({ storage: multer.memoryStorage() });

// Função para importar o arquivo
const importFile = async (req, res) => {
    try {
        const file = req.file;

        if (!file) {
            return res.status(400).json({ error: 'Arquivo não fornecido.' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(file.buffer);

        // Acessa a primeira planilha e seus valores
        const worksheet = workbook.worksheets[0];
        const data = worksheet.getSheetValues();

        // Salva o arquivo importado
        const outputPath = path.join(__dirname, 'imported_file.xlsx');
        await workbook.xlsx.writeFile(outputPath);

        res.json({ message: 'Arquivo importado com sucesso!', data });
    } catch (error) {
        console.error('Erro ao importar o arquivo:', error);
        res.status(500).json({ error: 'Erro ao importar o arquivo.' });
    }
};

module.exports = { upload, importFile };
