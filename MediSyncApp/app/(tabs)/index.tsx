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

  const [valueA1, setValueA1] = useState<string | null>(null); // Armazena o valor da célula A1

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
          // Ler o valor da célula A1
          const value = firstSheet['A1']?.v ?? '';

          console.log('Valor da célula A1:', value);

          // Armazenar o valor da célula A1 no estado
          setValueA1(value);

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

  const createExcelWithValues = async (rows: string[][]) => {
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

    const wbout = XLSX.write(newWorkbook, { type: 'base64', bookType: 'xlsx' });

    if (Platform.OS === 'web') {
      const blob = new Blob([Uint8Array.from(atob(wbout), c => c.charCodeAt(0))], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'new_excel.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url); // Clean up the URL object
      console.log('Arquivo Excel gerado e baixado no navegador.');
    } else {
      // Lógica para mobile (iOS e Android)
      const filename = 'new_excel.xlsx';
      const path = `${FileSystem.cacheDirectory}${filename}`;
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
    try {
      if (extractedRows.length === 0) {
        alert('Nenhum valor válido foi lido para as colunas de E a H.');
        return;
      }

      // Chamar a função para criar e baixar/compartilhar o Excel
      await createExcelWithValues(extractedRows);

      console.log('Processo de download/compartilhamento do Excel concluído.');
    } catch (error) {
      console.error('Erro ao processar o download/compartilhamento do Excel:', error);
      alert('Ocorreu um erro ao tentar processar o download/compartilhamento do Excel.');
    }
  };

  return (
    <View style={styles.container}>
      <View style={styles.footerContainer}>
        <Button theme="primary" label="Choose an Excel" onPress={pickExcelAsync} />
        <Button theme="primary" label="Download Excel" onPress={handleDownload} />
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