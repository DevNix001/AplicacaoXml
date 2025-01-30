import { Injectable } from '@angular/core';
import * as xml2js from 'xml2js';

@Injectable({
  providedIn: 'root'
})
export class XmlProcessorService {
  constructor() {}

  async processXmlFile(file: File): Promise<any> {
    const content = await this.readFileContent(file);
    return this.processXmlContent(content);
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
      const processedResult = this.processXmlNode(result);
      const flattenedData = this.flattenXmlData(processedResult);
      
      return { data: flattenedData };
    } catch (error) {
      console.error('Erro ao processar XML:', error);
      throw new Error('Erro ao processar arquivo XML');
    }
  }

  private processXmlNode(node: any): any {
    if (typeof node === 'string') return node;
    
    const result: any = {};
    
    if (node.$ && Object.keys(node.$).length > 0) {
      Object.assign(result, node.$);
    }
    
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

  private flattenXmlData(data: any[]): any[] {
    if (!Array.isArray(data)) {
      data = [data];
    }
    return data.map(item => this.flattenObject(item));
  }

  private flattenObject(obj: any, prefix = ''): any {
    const flattened: any = {};

    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        const value = obj[key];
        const newKey = prefix ? `${prefix}.${key}` : key;

        if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
          Object.assign(flattened, this.flattenObject(value, newKey));
        } else if (Array.isArray(value)) {
          flattened[key] = value.join(', ');
        } else {
          flattened[key] = value;
        }
      }
    }

    return flattened;
  }

  private readFileContent(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target?.result as string);
      reader.onerror = (e) => reject(e);
      reader.readAsText(file);
    });
  }
}
