// saveChangesHandler.js
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Função para salvar as alterações no arquivo importado
const saveChanges = async (req, res) => {
    try {
        const data = req.body.data;

        if (!data) {
            return res.status(400).json({ error: 'Dados não fornecidos para salvar.' });
        }

        // Caminho do arquivo importado
        const inputPath = path.join(__dirname, 'imported_file.xlsx');

        // Verifica se o arquivo existe e não está sendo usado
        if (!fs.existsSync(inputPath)) {
            return res.status(404).json({ error: 'Arquivo importado não encontrado.' });
        }

        // Lê o arquivo importado
        const readFile = async () => {
            return new Promise((resolve, reject) => {
                const checkInterval = 100; // Intervalo para verificar o arquivo
                let attempts = 10; // Número de tentativas para ler o arquivo

                const tryRead = async () => {
                    try {
                        const workbook = new ExcelJS.Workbook();
                        await workbook.xlsx.readFile(inputPath);
                        resolve(workbook);
                    } catch (error) {
                        if (attempts > 0) {
                            attempts--;
                            setTimeout(tryRead, checkInterval); // Tenta ler novamente após um intervalo
                        } else {
                            reject(error);
                        }
                    }
                };

                tryRead();
            });
        };

        const workbook = await readFile();

        // Verifica se a planilha "MP" existe
        const sheet = workbook.getWorksheet('MP');
        if (!sheet) {
            return res.status(404).json({ error: 'A planilha "MP" não foi encontrada.' });
        }

        // Atualiza a planilha com os novos dados, preservando a primeira linha
        data.forEach((row, rowIndex) => {
            row.forEach((cell, colIndex) => {
                const cellRef = sheet.getCell(rowIndex + 2, colIndex + 1); // Começa da linha 2 para preservar a primeira linha

                // Detecta se o valor é numérico e define o tipo apropriado
                if (!isNaN(cell) && cell !== '') {
                    cellRef.value = Number(cell); // Garante que o valor seja tratado como número
                    cellRef.numFmt = 'General'; // Define a formatação como número
                } else {
                    cellRef.value = cell;
                    cellRef.numFmt = '@'; // Define a formatação como texto
                }
            });
        });

        // Salva as alterações
        await workbook.xlsx.writeFile(inputPath);

        res.json({ message: 'Alterações salvas com sucesso!' });
    } catch (error) {
        console.error('Erro ao salvar as alterações:', error);
        res.status(500).json({ error: 'Erro ao salvar as alterações.' });
    }
};

module.exports = { saveChanges };
