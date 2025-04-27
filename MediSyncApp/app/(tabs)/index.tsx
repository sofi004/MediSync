import { useState } from 'react';
import { View, StyleSheet, Platform } from 'react-native';

import Button from '@/components/Button';
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import * as XLSX from 'xlsx';
import React from 'react';

export default function Index() {
  const [firstColumnValues, setFirstColumnValues] = useState<string[]>([]);
  const [extractedRows, setExtractedRows] = useState<string[][]>([]);

  const pickExcelAsync = async () => {
    let result = await DocumentPicker.getDocumentAsync({
      type: [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel', // .xls
        'text/csv', // .csv
      ],
      copyToCacheDirectory: true,
      multiple: false,
    });

    if (!result.canceled && result.assets?.[0]?.uri) {
      try {
        const fileUri = result.assets[0].uri;
        const response = await fetch(fileUri);
        const blob = await response.blob();

        const reader = new FileReader();
        reader.onloadend = async () => {
          const fileContent = reader.result?.toString().replace(/^data:.*;base64,/, '') || '';
          const workbook = XLSX.read(fileContent, { type: 'base64' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

          const rows: string[][] = [];
          let rowNumber = 1;

          while (firstSheet[`E${rowNumber}`]) {
            const e = firstSheet[`E${rowNumber}`]?.v?.toString() || '';
            const f = firstSheet[`F${rowNumber}`]?.v?.toString() || '';
            const g = firstSheet[`G${rowNumber}`]?.v?.toString() || '';
            const h = firstSheet[`H${rowNumber}`]?.v?.toString() || '';
            const i = firstSheet[`I${rowNumber}`]?.v?.toString() || '';
            const j = firstSheet[`J${rowNumber}`]?.v?.toString() || '';
            const k = firstSheet[`K${rowNumber}`]?.v?.toString() || '';
            const l = firstSheet[`L${rowNumber}`]?.v?.toString() || '';
            const m = firstSheet[`M${rowNumber}`]?.v?.toString() || '';
            const n = firstSheet[`N${rowNumber}`]?.v?.toString() || '';
            const o = firstSheet[`O${rowNumber}`]?.v?.toString() || '';
            const p = firstSheet[`P${rowNumber}`]?.v?.toString() || '';
            const q = firstSheet[`Q${rowNumber}`]?.v?.toString() || '';
            const r = firstSheet[`R${rowNumber}`]?.v?.toString() || '';
            const s = firstSheet[`S${rowNumber}`]?.v?.toString() || '';
            const t = firstSheet[`T${rowNumber}`]?.v?.toString() || '';
            const u = firstSheet[`U${rowNumber}`]?.v?.toString() || '';
            const v = firstSheet[`V${rowNumber}`]?.v?.toString() || '';
            const w = firstSheet[`W${rowNumber}`]?.v?.toString() || '';
            const x = firstSheet[`X${rowNumber}`]?.v?.toString() || '';
            const y = firstSheet[`Y${rowNumber}`]?.v?.toString() || '';
            const z = firstSheet[`Z${rowNumber}`]?.v?.toString() || '';
            rows.push([e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z]);
            rowNumber++;
          }

          setExtractedRows(rows);

          const values: string[] = [];
          let row = 1;
          let hasDuplicates = false;
          const seenValues = new Set<string>();

          while (firstSheet[`A${row}`]) {
            const cellValue = firstSheet[`A${row}`].v?.toString() || '';
            if (seenValues.has(cellValue)) {
              alert(`Valor duplicado encontrado na célula A${row}: ${cellValue}`);
              hasDuplicates = true;
              break; // Interrompe a leitura ao encontrar a primeira duplicata
            }
            seenValues.add(cellValue);
            values.push(cellValue);
            console.log(`Valor na célula A${row}:`, cellValue);
            row++;
          }

          if (hasDuplicates) {
            alert('O arquivo contém valores duplicados na primeira coluna. A operação foi interrompida.');
            setFirstColumnValues([]); // Não atualiza o estado se houver duplicatas
          } else if (values.length > 0) {
            console.log('Valores lidos da primeira coluna (sem duplicatas):', values);
            setFirstColumnValues(values);
          } else {
            alert('Nenhum valor encontrado na primeira coluna.');
            setFirstColumnValues([]);
          }
        };
        reader.readAsDataURL(blob);
      } catch (error) {
        console.error('Erro ao processar o ficheiro:', error);

        let errorMessage = 'Erro ao processar o ficheiro. ';
        if (error instanceof TypeError) {
          errorMessage += 'Parece que o arquivo selecionado não é válido ou está corrompido.';
        } else if (error instanceof SyntaxError) {
          errorMessage += 'Houve um problema ao interpretar o conteúdo do arquivo.';
        } else if (error instanceof Error) {
          errorMessage += `Detalhes: ${error.message || 'Erro desconhecido.'}`;
        } else {
          errorMessage += 'Erro desconhecido.';
        }

        alert(errorMessage);
      }
    } else {
      alert('Você não selecionou nenhum arquivo Excel.');
    }
  };

  const createExcelWithValues = async (templateRows: string[][]) => {
    if (!templateRows || templateRows.length === 0) {
      alert('Nenhum dado de semana encontrado para repetir.');
      return;
    }

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheetData: any[][] = [];

    let currentDate = new Date('2024-10-01T00:00:00');
    const endDate = new Date('2025-12-31T23:59:59');
    const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
    const dayColumns = ['', 'B', '', 'E', '', 'H', '', 'K', '', 'N', '', 'Q', '', 'T', '', 'W', '', 'Z']; // Colunas onde os dias começam

    while (currentDate <= endDate) {
      const weekHeader = Array(26).fill(''); // Inicializa um array com o tamanho esperado de colunas
      const firstDayOfWeek = new Date(currentDate);
      // Encontra o primeiro dia da semana (Segunda-feira)
      while (firstDayOfWeek.getDay() !== 1) {
        firstDayOfWeek.setDate(firstDayOfWeek.getDate() - 1);
      }

      for (let i = 0; i < 7; i++) {
        const day = new Date(firstDayOfWeek);
        day.setDate(firstDayOfWeek.getDate() + i);

        if (day <= endDate) {
          const dayName = daysOfWeek[day.getDay()];
          const dateString = day.toLocaleDateString();
          let columnIndex = -1;

          if (dayName === 'Segunda') columnIndex = 1; // Column B
          else if (dayName === 'Terça') columnIndex = 4; // Column E
          else if (dayName === 'Quarta') columnIndex = 7; // Column H
          else if (dayName === 'Quinta') columnIndex = 10; // Column K
          else if (dayName === 'Sexta') columnIndex = 13; // Column N
          else if (dayName === 'Sábado') columnIndex = 16; // Column Q
          else if (dayName === 'Domingo') columnIndex = 19; // Column T

          if (columnIndex !== -1 && day.getMonth() === firstDayOfWeek.getMonth()) {
            weekHeader[columnIndex] = day.toLocaleDateString('pt-PT', { day: 'numeric', month: 'numeric', year: 'numeric' });
          } else if (columnIndex !== -1) {
            weekHeader[columnIndex] = day.toLocaleDateString('pt-PT', { day: 'numeric', month: 'numeric', year: 'numeric' });
          }
        }
      }
      newWorksheetData.push(weekHeader);

      // Adicionar as linhas da semana do template
      templateRows.forEach(row => {
        newWorksheetData.push([...row]);
      });

      // Avança para a próxima semana
      currentDate.setDate(currentDate.getDate() + 7);
    }

    const newWorksheet = XLSX.utils.aoa_to_sheet(newWorksheetData);

    // Definir a largura das colunas (opcional, mas melhora a visualização)
    // Definir a largura das colunas (opcional, mas melhora a visualização)
    const wscols = [
      { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 15 },
      { wch: 10 }, { wch: 10 }, { wch: 10 },
      { wch: 10 }, { wch: 10 }, { wch: 10 },
      { wch: 10 }, { wch: 10 }, { wch: 10 },
      { wch: 10 }, { wch: 10 }, { wch: 10 },
      { wch: 10 }, { wch: 10 }, { wch: 10 },
      { wch: 10 }, { wch: 10 }, { wch: 10 },
      { wch: 10 }, { wch: 10 }, { wch: 10 },
    ];
    newWorksheet['!cols'] = wscols;

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Calendário');

    const wbout = XLSX.write(newWorkbook, { type: 'base64', bookType: 'xlsx' });

    if (Platform.OS === 'web') {
      const blob = new Blob([Uint8Array.from(atob(wbout), c => c.charCodeAt(0))], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'calendario_2024_2025.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url); // Clean up the URL object
      console.log('Arquivo Excel gerado e baixado no navegador.');
    } else {
      // Lógica para mobile (iOS e Android)
      const filename = 'calendario_2024_2025.xlsx';
      const path = `<span class="math-inline">\{FileSystem\.cacheDirectory\}</span>{filename}`;
      try {
        await FileSystem.writeAsStringAsync(path, wbout, { encoding: FileSystem.EncodingType.Base64 });
        console.log('Novo ficheiro criado em:', path);
        if (await Sharing.isAvailableAsync()) {
          await Sharing.shareAsync(path, {
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            dialogTitle: 'Compartilhar arquivo Excel',
          });
          console.log('Arquivo Excel compartilhado com sucesso.');
        } else {
          alert('Compartilhamento não disponível neste dispositivo.');
        }
      } catch (error) {
        console.error('Erro ao criar ou compartilhar o arquivo Excel:', error);
        alert('Ocorreu um erro ao criar ou compartilhar o arquivo Excel.');
      }
    }
  };

  const handleDownload = async () => {
    if (extractedRows.length === 0) {
      alert('Por favor, carregue um arquivo Excel primeiro para obter a estrutura das semanas.');
      return;
    }

    // Filtrar as linhas que contêm dados de morning/afternoon para o template da semana
    const weekTemplate: string[][] = [];
    for (let i = 0; i < extractedRows.length; i++) {
      weekTemplate.push(extractedRows[i]);
    }

    if (weekTemplate.length === 0) {
      alert('Não foram encontradas semanas válidas no arquivo carregado (precisa de linhas com "morning" e "afternoon").');
      return;
    }

    await createExcelWithValues(weekTemplate);
  };

  return (
    <View style={styles.container}>
      <View style={styles.footerContainer}>
        <Button theme="primary" label="Escolher Excel" onPress={pickExcelAsync} />
        <Button theme="primary" label="Baixar Excel" onPress={handleDownload} />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#25292e',
    alignItems: 'center',
    justifyContent: 'center',
  },
  footerContainer: {
    gap: 16,
    alignItems: 'center',
  },
});