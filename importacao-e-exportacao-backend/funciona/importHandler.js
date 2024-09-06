const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const { saveImportedFile } = require('./fileUtils'); // Certifique-se de importar a função corretamente

// Configuração do multer para lidar com uploads de arquivos
const upload = multer({ storage: multer.memoryStorage() });

// Função para importar o arquivo
const importFile = (req, res) => {
    try {
        // Verifique se o arquivo foi recebido
        console.log('Arquivo recebido:', req.file);

        const file = req.file;

        if (!file) {
            return res.status(400).json({ error: 'Arquivo não fornecido.' });
        }

        // Lê o buffer do arquivo como uma planilha
        const workbook = XLSX.read(file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0]; // Pega a primeira planilha
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Salva o arquivo importado
        const outputPath = path.join(__dirname, 'imported_file.xlsx');
        saveImportedFile(workbook, outputPath); // Atualize aqui

        res.json({ message: 'Arquivo importado com sucesso!', data });
    } catch (error) {
        console.error('Erro ao importar o arquivo:', error);
        res.status(500).json({ error: 'Erro ao importar o arquivo.' });
    }
};

module.exports = { upload, importFile };
