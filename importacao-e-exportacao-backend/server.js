// server.js
const express = require('express');
const cors = require('cors');
const { upload, importFile } = require('./importHandler');
const { exportFile } = require('./exportHandler');
const { saveChanges } = require('./saveChangesHandler');
const { exportPdf } = require('./pdfExportHandler'); // Importa o novo handler
const path = require('path');

const app = express();
const port = 3001;

app.use(cors());
app.use(express.json({ limit: '50mb' })); // Permite o upload de arquivos grandes
app.use(express.urlencoded({ extended: true }));

// Rota para importação de arquivos
app.post('/import', upload.single('file'), importFile);

// Rota para exportação de arquivos em Excel
app.post('/export/xlsx', exportFile);

// Rota para exportação de arquivos em PDF
app.post('/export/pdf', exportPdf);

// Rota para salvar alterações no arquivo importado
app.post('/save', saveChanges);

app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});
