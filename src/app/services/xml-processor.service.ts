import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import parse from 'xml-parser';

interface AdicaoData {
  [key: string]: string;
}

@Injectable({
  providedIn: 'root'
})
export class XmlProcessorService {

  constructor() { }

  async processXmlFile(file: File): Promise<{ columns: string[], data: AdicaoData[] }> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const xmlContent = e.target?.result as string;
          const data = this.parseXml(xmlContent);

          const columns = new Set<string>();
          data.forEach(item => {
            Object.keys(item).forEach(key => columns.add(key));
          });

          resolve({
            columns: Array.from(columns),
            data
          });
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsText(file);
    });
  }

  parseXml(xmlString: string): AdicaoData[] {
    const xml = parse(xmlString);
    const declarations = xml.root.children;
    const data: AdicaoData[] = [];
  
    declarations.forEach((declaration: any) => {
      if (declaration.name === 'declaracaoImportacao') {
        declaration.children.forEach((child: any) => {
          if (child.name === 'adicao') {
            const adicaoData = this.extractAdicaoData(child);
            data.push(adicaoData);
          }
        });
      }
    });
    return data;
  }

  private extractAdicaoData(adicaoNode: any): AdicaoData {
    const adicaoData: AdicaoData = {};
    adicaoNode.children.forEach((item: any) => {
      adicaoData[item.name] = item.content;
    });
    return adicaoData;
  }

  async exportToExcel(data: AdicaoData[]): Promise<void> {
    try {
      // Cria a planilha sem opções especiais
      const worksheet = XLSX.utils.json_to_sheet(data);

      // Ajusta a largura das colunas baseado no conteúdo
      const maxWidth = Object.keys(data[0] || {}).reduce((acc, key) => {
        const maxLength = Math.max(
          key.length,
          ...data.map(row => (row[key]?.toString() || '').length)
        );
        acc[key] = maxLength + 2; // +2 para padding
        return acc;
      }, {} as { [key: string]: number });

      worksheet['!cols'] = Object.keys(maxWidth).map(key => ({
        wch: maxWidth[key]
      }));

      // Garante que os zeros sejam mantidos convertendo para texto
      if (data.length > 0) {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: C })];
            if (cell && cell.v !== undefined && cell.v !== null) {
              cell.t = 's'; // Define o tipo como string
              cell.v = cell.v.toString(); // Converte para string
            }
          }
        }
      }

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Dados Combinados');
      
      XLSX.writeFile(workbook, 'dados_exportados.xlsx');
    } catch (error) {
      console.error('Erro ao exportar para Excel:', error);
      throw error;
    }
  }
}
