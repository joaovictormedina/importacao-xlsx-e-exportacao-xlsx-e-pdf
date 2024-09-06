import React, { useState } from 'react';
import { Table, Container, Input, Text, Button, MantineProvider } from '@mantine/core';
import * as XLSX from 'xlsx';
import './App.css';  // Certifique-se de que este arquivo está incluído

function App() {
  const [tableData, setTableData] = useState(Array(10).fill(Array(36).fill(''))); // 10 linhas e 36 colunas iniciais

  // Função para enviar os dados importados para o servidor
  const uploadData = (file) => {
    const formData = new FormData();
    formData.append('file', file);

    fetch('http://localhost:3001/import', {
      method: 'POST',
      body: formData,
    })
      .then(response => response.json())
      .then(result => console.log(result))
      .catch(error => console.error('Erro ao enviar dados para o servidor:', error));
  };

  // Função para exportar os dados da tabela
  const exportData = () => {
    if (!tableData || tableData.length === 0) {
      console.error('Nenhum dado para exportar.');
      return;
    }

    fetch('http://localhost:3001/export/xlsx', {
      method: 'POST',
      body: JSON.stringify({ data: tableData }), // Envia os dados da tabela para exportação
      headers: {
        'Content-Type': 'application/json',
      },
    })
      .then(response => response.blob())
      .then(blob => {
        const url = window.URL.createObjectURL(new Blob([blob]));
        const a = document.createElement('a');
        a.href = url;
        a.download = 'exported_data.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
      })
      .catch(error => console.error('Erro ao exportar dados:', error));
  };

  // Função para manipular a mudança de arquivo
  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();

      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Verifica se a planilha "MP" existe
        if (workbook.SheetNames.includes("MP")) {
          const sheet = workbook.Sheets["MP"];
          const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: true });

          // Define o número de colunas da tabela
          const numColumns = 36;

          // Preenche os dados na tabela
          const formattedData = jsonData.slice(1).map(row => {
            const newRow = Array(numColumns).fill(''); // Preenche com células vazias
            row.forEach((cell, index) => {
              if (index < numColumns) {
                newRow[index] = cell;
              }
            });
            return newRow;
          });

          // Ajusta o tamanho da tabela para preencher a quantidade de linhas
          const fullData = [...tableData];
          for (let i = 0; i < formattedData.length; i++) {
            fullData[i] = formattedData[i];
          }

          setTableData(fullData);

          // Envia os dados importados para o servidor
          uploadData(file);
        } else {
          alert('A planilha "MP" não foi encontrada na pasta de trabalho.');
        }
      };

      reader.readAsArrayBuffer(file);
    }
  };

  // Função para manipular a mudança de célula
  const handleCellChange = (rowIndex, colIndex, value) => {
    setTableData(prevData => {
      const newData = [...prevData];
      const newRow = [...newData[rowIndex]];
      newRow[colIndex] = value;
      newData[rowIndex] = newRow;
      return newData;
    });
  };

  return (
    <MantineProvider withGlobalStyles withNormalizeCSS>
      <Container>
        <Text size="xl" weight={500} align="center" style={{ marginBottom: '20px' }}>
          Importar Excel
        </Text>
        <Input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileChange}
          style={{ marginBottom: '20px' }}
        />
        <div className="table-container">
          <Table highlightOnHover className="editable-table">
            <thead>
              <tr>
                <th>Item</th>
                <th>P/N</th>
                <th>Descrição</th>
                <th>Peso bruto</th>
                <th>Peso Líq</th>
                <th>% Resina</th>
                <th>Resina</th>
                <th>% Master</th>
                <th>Master</th>
                <th>Qtd/Cx</th>
                <th>Caixa</th>
                <th>Preço Caixa</th>
                <th>Componente</th>
                <th>Qtd Comp</th>
                <th>Saco plast</th>
                <th>Qtd saco</th>
                <th>Molde</th>
                <th>Preço Unit</th>
                <th>MAQ</th>
                <th>Ciclo (s)</th>
                <th>Hora Máq</th>
                <th>Cavid</th>
                <th>Perda</th>
                <th>Comentário</th>
                <th>Frete</th>
                <th>Cx/Caminhão</th>
                <th>Reutilização</th>
                <th>Processo</th>
                <th>Pc/h</th>
                <th>HH</th>
                <th>Lucro</th>
                <th>Tx Fin</th>
                <th>Ref. Valor HM</th>
                <th>Qtd/Mês</th>
                <th>H/Mês</th>
              </tr>
            </thead>
            <tbody>
              {tableData.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, colIndex) => (
                    <td key={colIndex}>
                      <input
                        type="text"
                        value={cell}
                        onChange={(e) => handleCellChange(rowIndex, colIndex, e.target.value)}
                        className="editable-cell"
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </Table>
        </div>
        <Button
          onClick={exportData}
          style={{ marginTop: '20px' }}
        >
          Exportar para Excel
        </Button>
      </Container>
    </MantineProvider>
  );
}

export default App;
