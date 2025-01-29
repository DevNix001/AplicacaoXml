import { Injectable } from '@angular/core';
import * as xml2js from 'xml2js';
import * as XLSX from 'xlsx';

@Injectable({
  providedIn: 'root'
})
export class XmlService {
  constructor() { }

  parseXmlFile(file: File): Promise<any> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const xml = e.target?.result as string;
        xml2js.parseString(xml, (err: any, result: any) => {
          if (err) {
            reject(err);
          } else {
            resolve(result);
          }
        });
      };
      reader.onerror = (e) => reject(e);
      reader.readAsText(file);
    });
  }

  exportToExcel(data: any[], fileName: string): void {
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Declarações');
    
    const colWidths = this.calculateColumnWidths(data);
    ws['!cols'] = colWidths;
    
    XLSX.writeFile(wb, `${fileName}.xlsx`);
  }

  exportMultipleToExcel(files: { name: string, data: any[] }[]): void {
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    
    files.forEach(file => {
      const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(file.data);
      const colWidths = this.calculateColumnWidths(file.data);
      ws['!cols'] = colWidths;
      XLSX.utils.book_append_sheet(wb, ws, file.name.substring(0, 31)); // Excel limita nome da aba a 31 caracteres
    });
    
    XLSX.writeFile(wb, `Declaracoes_Importacao_${new Date().getTime()}.xlsx`);
  }

  flattenXmlData(xmlData: any): any[] {
    const flatData: any[] = [];
    try {
      if (xmlData.ListaDeclaracoes && xmlData.ListaDeclaracoes.declaracaoImportacao) {
        const declaracoes = Array.isArray(xmlData.ListaDeclaracoes.declaracaoImportacao) 
          ? xmlData.ListaDeclaracoes.declaracaoImportacao 
          : [xmlData.ListaDeclaracoes.declaracaoImportacao];

        declaracoes.forEach((declaracao: any) => {
          if (declaracao.adicao) {
            const adicoes = Array.isArray(declaracao.adicao) 
              ? declaracao.adicao 
              : [declaracao.adicao];

            adicoes.forEach((adicao: any) => {
              const flatAdicao: any = {};
              
              Object.keys(adicao).forEach(key => {
                if (typeof adicao[key] === 'object' && adicao[key] !== null) {
                  Object.keys(adicao[key]).forEach(subKey => {
                    flatAdicao[`${key}_${subKey}`] = this.cleanValue(adicao[key][subKey]);
                  });
                } else {
                  flatAdicao[key] = this.cleanValue(adicao[key]);
                }
              });
              
              flatData.push(flatAdicao);
            });
          }
        });
      }
    } catch (error) {
      console.error('Erro ao processar XML:', error);
    }
    return flatData;
  }

  private cleanValue(value: any): string {
    if (Array.isArray(value)) {
      return value[0]?.toString() || '';
    }
    return value?.toString().replace(/#x20;/g, ' ').trim() || '';
  }

  private calculateColumnWidths(data: any[]): any[] {
    const colWidths: any[] = [];
    if (data.length > 0) {
      const headers = Object.keys(data[0]);
      headers.forEach(header => {
        let maxLength = header.length;
        data.forEach(row => {
          const cellLength = (row[header]?.toString() || '').length;
          maxLength = Math.max(maxLength, cellLength);
        });
        colWidths.push({ wch: Math.min(maxLength + 2, 50) });
      });
    }
    return colWidths;
  }
}
