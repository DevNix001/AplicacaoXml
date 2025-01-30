import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import parse from 'xml-parser';
import * as xml2js from 'xml2js';

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

  async processXmlContent(xmlContent: string): Promise<any> {
    try {
      const parser = new xml2js.Parser({
        explicitArray: false,
        trim: true,
        explicitRoot: false,
        mergeAttrs: true,
        attrNameProcessors: [(name: string) => name.toLowerCase()],
        tagNameProcessors: [(name: string) => name.toLowerCase()]
      });

      const result = await parser.parseStringPromise(xmlContent);
      
      // Processa o resultado para garantir uma estrutura consistente
      const processedResult = this.processXmlNode(result);
      
      // Achata a estrutura
      const flattenedData = this.flattenXmlData(processedResult);
      
      return { data: flattenedData };
    } catch (error) {
      console.error('Erro ao processar XML:', error);
      throw new Error('Erro ao processar arquivo XML');
    }
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

  private flattenXmlData(data: any): any[] {
    if (!data) return [];

    // Se data for um array, processa cada item
    if (Array.isArray(data)) {
      return data.map(item => this.flattenObject(item));
    }

    // Se data não for um array mas contiver um array interno
    for (const key in data) {
      if (Array.isArray(data[key])) {
        return data[key].map((item: any) => this.flattenObject(item));
      }
    }

    // Se não encontrar array, retorna o objeto único em um array
    return [this.flattenObject(data)];
  }

  private flattenObject(obj: any, prefix = ''): any {
    const flattened: any = {};

    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        const value = obj[key];
        const newKey = prefix ? `${prefix}.${key}` : key;

        if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
          // Se for um objeto, achata recursivamente
          Object.assign(flattened, this.flattenObject(value, newKey));
        } else if (Array.isArray(value)) {
          // Se for um array, junta os valores com vírgula
          flattened[key] = value.join(', ');
        } else {
          // Para valores simples, mantém como está
          flattened[key] = value;
        }
      }
    }

    return flattened;
  }

  private processXmlNode(node: any): any {
    if (typeof node === 'string') return node;
    
    const result: any = {};
    
    // Processa atributos
    if (node.$ && Object.keys(node.$).length > 0) {
      Object.assign(result, node.$);
    }
    
    // Processa nós filhos
    for (const key in node) {
      if (key !== '$') {
        const value = node[key];
        if (Array.isArray(value)) {
          result[key] = value.map(item => this.processXmlNode(item));
        } else if (typeof value === 'object') {
          result[key] = this.processXmlNode(value);
        } else {
          result[key] = value;
        }
      }
    }
    
    return result;
  }

  private async readFileContent(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target?.result as string);
      reader.onerror = (e) => reject(e);
      reader.readAsText(file);
    });
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
