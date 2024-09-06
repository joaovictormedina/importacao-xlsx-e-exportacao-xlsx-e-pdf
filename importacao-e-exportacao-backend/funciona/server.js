// server.js
const express = require('express');
const cors = require('cors');
const { upload, importFile } = require('./importHandler');
const { exportFile } = require('./exportHandler');

const app = express();
const port = 3001;

app.use(cors());
app.use(express.json({ limit: '50mb' })); // Permite o upload de arquivos grandes
app.use(express.urlencoded({ extended: true }));

// Rota para importação de arquivos
app.post('/import', upload.single('file'), importFile);

// Rota para exportação de arquivos
app.post('/export/xlsx', exportFile);

app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});
