import { Component, ElementRef, HostListener, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { MatCardModule } from '@angular/material/card';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatTableModule } from '@angular/material/table';
import { MatTooltipModule } from '@angular/material/tooltip';
import { DragDropModule, CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import * as xml2js from 'xml2js';

interface Column {
  key: string;
  displayName: string;
  selected: boolean;
  visible: boolean;
}

interface ImportedFile {
  name: string;
  data: any[];
  columns: Column[];
  columnSearchText: string;
}

@Component({
  selector: 'app-xml-import',
  templateUrl: './xml-import.component.html',
  styleUrls: ['./xml-import.component.scss'],
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatCardModule,
    MatButtonModule,
    MatIconModule,
    MatCheckboxModule,
    MatInputModule,
    MatFormFieldModule,
    MatTableModule,
    MatTooltipModule,
    DragDropModule
  ]
})
export class XmlImportComponent {
  @ViewChild('top') topElement!: ElementRef;
  importedFiles: ImportedFile[] = [];
  showBackToTop = false;

  @HostListener('window:scroll')
  onWindowScroll() {
    this.showBackToTop = window.scrollY > 300;
  }

  scrollToTop() {
    this.topElement?.nativeElement.scrollIntoView({ behavior: 'smooth' });
  }

  async onFileSelected(event: any) {
    const files = event.target.files;
    if (files) {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const fileContent = await this.readFileContent(file);
        const jsonData = await this.convertXmlToJson(fileContent);
        const columns = this.extractColumns(jsonData);
        
        this.importedFiles.push({
          name: file.name,
          data: jsonData,
          columns: columns,
          columnSearchText: ''
        });
      }
    }
  }

  private readFileContent(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target?.result as string);
      reader.onerror = (e) => reject(e);
      reader.readAsText(file);
    });
  }

  private convertXmlToJson(xmlContent: string): Promise<any[]> {
    return new Promise((resolve, reject) => {
      const parser = new xml2js.Parser({ 
        explicitArray: false,
        mergeAttrs: true,
        explicitRoot: false
      });
      
      parser.parseString(xmlContent, (err: any, result: any) => {
        if (err) {
          reject(err);
          return;
        }
        
        const rows = result.row || [];
        resolve(Array.isArray(rows) ? rows : [rows]);
      });
    });
  }

  private extractColumns(data: any[]): Column[] {
    const columnSet = new Set<string>();
    data.forEach(item => {
      Object.keys(item).forEach(key => columnSet.add(key));
    });

    return Array.from(columnSet).map(key => ({
      key,
      displayName: key,
      selected: true,
      visible: true
    }));
  }

  getVisibleColumns(file: ImportedFile): Column[] {
    return file.columns.filter(column => {
      const searchMatch = !file.columnSearchText || 
        column.displayName.toLowerCase().includes(file.columnSearchText.toLowerCase());
      return column.visible && searchMatch;
    });
  }

  toggleColumnSelection(column: Column, file: ImportedFile) {
    column.selected = !column.selected;
  }

  toggleAllColumns(checked: boolean, file: ImportedFile) {
    file.columns.forEach(column => {
      if (column.visible) {
        column.selected = checked;
      }
    });
  }

  areAllColumnsSelected(file: ImportedFile): boolean {
    const visibleColumns = this.getVisibleColumns(file);
    return visibleColumns.length > 0 && visibleColumns.every(column => column.selected);
  }

  areSomeColumnsSelected(file: ImportedFile): boolean {
    const visibleColumns = this.getVisibleColumns(file);
    return visibleColumns.some(column => column.selected) && !this.areAllColumnsSelected(file);
  }

  getVisibleColumnKeys(file: ImportedFile): string[] {
    return this.getVisibleColumns(file).map(column => column.key);
  }

  removeFile(index: number) {
    this.importedFiles.splice(index, 1);
  }

  onDrop(event: CdkDragDrop<string[]>, file: ImportedFile) {
    moveItemInArray(file.columns, event.previousIndex, event.currentIndex);
  }

  canExport(): boolean {
    return this.importedFiles.some(file => 
      file.columns.some(column => column.selected)
    );
  }

  async exportAllToExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados Combinados');
    
    // Mapa para armazenar todas as colunas únicas
    const uniqueColumns = new Map<string, Set<string>>();
    
    // Primeiro passo: Coletar todas as colunas selecionadas e seus possíveis nomes diferentes
    this.importedFiles.forEach(file => {
      file.columns.forEach(column => {
        if (column.selected) {
          // Normaliza o nome da coluna (remove espaços extras, converte para minúsculo)
          const normalizedName = this.normalizeColumnName(column.displayName);
          
          if (!uniqueColumns.has(normalizedName)) {
            uniqueColumns.set(normalizedName, new Set());
          }
          uniqueColumns.get(normalizedName)?.add(column.key);
        }
      });
    });

    // Converter o mapa de colunas únicas em um array de cabeçalhos
    const headers = Array.from(uniqueColumns.keys());
    worksheet.addRow(headers);

    // Adicionar dados de cada arquivo
    this.importedFiles.forEach(file => {
      file.data.forEach(row => {
        const rowData = headers.map(header => {
          // Procura por qualquer coluna que corresponda ao cabeçalho normalizado
          const possibleKeys = uniqueColumns.get(header) || new Set();
          for (const key of possibleKeys) {
            if (row[key] !== undefined && row[key] !== '') {
              return row[key];
            }
          }
          return '';
        });
        worksheet.addRow(rowData);
      });
    });

    // Ajustar largura das colunas
    worksheet.columns.forEach(column => {
      column.width = 20;
    });

    // Estilizar cabeçalhos
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E0E0' }
    };

    // Gerar arquivo
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'dados_combinados.xlsx');
  }

  private normalizeColumnName(name: string): string {
    // Remove espaços extras, converte para minúsculo e remove acentos
    return name
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, ' ');
  }
}
