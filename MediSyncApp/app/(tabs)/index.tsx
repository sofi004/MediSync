import { useState } from 'react';
import { View, StyleSheet } from 'react-native';

import Button from '@/components/Button';
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import * as XLSX from 'xlsx';
import React from 'react';

export default function Index() {
  const [valueA1, setValueA1] = useState<string | null>(null); // Armazena o valor da célula A1

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

        // Ler o arquivo usando fetch (compatível com web)
        const response = await fetch(fileUri);
        const blob = await response.blob();

        // Converter o blob para base64 usando FileReader
        const reader = new FileReader();
        reader.onloadend = async () => {
          const fileContent = reader.result?.toString().replace(/^data:.*;base64,/, '') || '';

          // Processar o conteúdo com XLSX
          const workbook = XLSX.read(fileContent, { type: 'base64' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const value = firstSheet['A1']?.v ?? '';

          console.log('Valor da célula A1:', value);

          // Armazenar o valor da célula A1 no estado
          setValueA1(value);
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

  const createExcelWithValueA1 = async (value: string) => {
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet([[value]]);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

    // Gerar o arquivo Excel em base64
    const wbout = XLSX.write(newWorkbook, {
      type: 'base64',
      bookType: 'xlsx',
    });

    // Verificar se está rodando na web
    if (typeof window !== 'undefined') {
      // Criar um link de download dinâmico
      const blob = new Blob([Uint8Array.from(atob(wbout), c => c.charCodeAt(0))], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);

      // Criar um link para download
      const link = document.createElement('a');
      link.href = url;
      link.download = 'new_excel.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      console.log('Arquivo Excel gerado e pronto para download.');
    } else {
      // Caso esteja em um dispositivo móvel, usar expo-file-system
      const path = FileSystem.cacheDirectory + 'new_excel.xlsx';
      await FileSystem.writeAsStringAsync(path, wbout, {
        encoding: FileSystem.EncodingType.Base64,
      });

      console.log('Novo ficheiro criado em:', path);
    }
  };

  const handleDownload = async () => {
    try {
      if (!valueA1) {
        alert('Nenhum valor foi processado. Por favor, escolha um arquivo Excel primeiro.');
        return;
      }

      // Chamar a função para criar e baixar o Excel
      await createExcelWithValueA1(valueA1);

      console.log('Download do Excel concluído.');
    } catch (error) {
      console.error('Erro ao fazer o download do Excel:', error);
      alert('Ocorreu um erro ao tentar fazer o download do Excel.');
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