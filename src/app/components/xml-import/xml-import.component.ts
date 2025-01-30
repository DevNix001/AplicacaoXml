import { Component, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { MatButtonModule } from '@angular/material/button';
import { MatCardModule } from '@angular/material/card';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatIconModule } from '@angular/material/icon';
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatTableModule } from '@angular/material/table';
import { MatTooltipModule } from '@angular/material/tooltip';
import { MatSnackBar, MatSnackBarModule } from '@angular/material/snack-bar';
import { DragDropModule, CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import { XmlProcessorService } from '../../services/xml-processor.service';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

interface Column {
  key: string;
  displayName: string;
  selected: boolean;
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
    MatButtonModule,
    MatCardModule,
    MatCheckboxModule,
    MatIconModule,
    MatInputModule,
    MatFormFieldModule,
    MatTableModule,
    MatTooltipModule,
    DragDropModule,
    MatSnackBarModule
  ]
})
export class XmlImportComponent {
  @ViewChild('top') topElement!: ElementRef;
  importedFiles: ImportedFile[] = [];
  showBackToTop = false;

  constructor(
    private xmlProcessor: XmlProcessorService,
    private snackBar: MatSnackBar
  ) {
    window.addEventListener('scroll', () => {
      this.showBackToTop = window.scrollY > 300;
    });
  }

  async onFileSelected(event: any): Promise<void> {
    const files: FileList = event.target?.files || event.dataTransfer?.files;
    if (!files) return;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (file.type !== 'text/xml' && !file.name.endsWith('.xml')) continue;

      try {
        const result = await this.xmlProcessor.processXmlFile(file);
        if (result.data.length > 0) {
          // Verifica se já existem arquivos importados para manter a consistência das colunas
          let columns: Column[];
          if (this.importedFiles.length > 0) {
            // Usa as colunas do primeiro arquivo como referência
            const existingColumns = this.importedFiles[0].columns;
            // Mantém as colunas existentes e adiciona novas se houver
            const newColumnKeys = Object.keys(result.data[0]);
            columns = existingColumns.map(col => ({...col})); // Clone as colunas existentes
            
            // Adiciona novas colunas que não existiam antes
            newColumnKeys.forEach(key => {
              const normalizedKey = this.normalizeColumnName(key);
              if (!columns.find(col => this.normalizeColumnName(col.key) === normalizedKey)) {
                columns.push({
                  key,
                  displayName: key,
                  selected: true
                });
              }
            });
          } else {
            // Se for o primeiro arquivo, cria as colunas normalmente
            columns = Object.keys(result.data[0]).map(key => ({
              key,
              displayName: key,
              selected: true
            }));
          }

          // Normaliza os dados para usar as mesmas chaves quando as colunas têm o mesmo nome
          const normalizedData = result.data.map(item => {
            const normalizedItem: any = {};
            Object.entries(item).forEach(([key, value]) => {
              const normalizedKey = this.normalizeColumnName(key);
              const targetColumn = columns.find(col => this.normalizeColumnName(col.key) === normalizedKey);
              if (targetColumn) {
                normalizedItem[targetColumn.key] = value;
              }
            });
            return normalizedItem;
          });

          // Atualiza as colunas em todos os arquivos importados para manter consistência
          if (this.importedFiles.length > 0) {
            this.importedFiles.forEach(importedFile => {
              importedFile.columns = columns.map(col => ({...col}));
            });
          }

          this.importedFiles.push({
            name: file.name,
            data: normalizedData,
            columns: columns.map(col => ({...col})),
            columnSearchText: ''
          });
        }
      } catch (error) {
        console.error('Erro ao processar arquivo XML:', error);
        this.showErrorMessage('Não foi possível processar o arquivo XML. Verifique se o formato está correto.');
      }
    }

    if (event.target) event.target.value = '';
  }

  removeFile(index: number): void {
    this.importedFiles.splice(index, 1);
  }

  getVisibleColumns(file: ImportedFile): Column[] {
    return file.columns.filter(column =>
      !file.columnSearchText ||
      column.displayName.toLowerCase().includes(file.columnSearchText.toLowerCase())
    );
  }

  getVisibleColumnKeys(file: ImportedFile): string[] {
    return this.getVisibleColumns(file).map(col => col.key);
  }

  toggleAllColumns(checked: boolean, file: ImportedFile): void {
    const visibleColumns = this.getVisibleColumns(file);
    visibleColumns.forEach(column => column.selected = checked);
    
    // Atualiza o mesmo estado em todas as instâncias da coluna em outros arquivos
    const columnKeys = visibleColumns.map(col => this.normalizeColumnName(col.key));
    this.importedFiles.forEach(otherFile => {
      otherFile.columns.forEach(col => {
        if (columnKeys.includes(this.normalizeColumnName(col.key))) {
          col.selected = checked;
        }
      });
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

  toggleColumnSelection(column: Column, file: ImportedFile): void {
    column.selected = !column.selected;
  }

  onColumnDrop(event: CdkDragDrop<string[]>, file: ImportedFile): void {
    if (event.previousIndex !== event.currentIndex) {
      moveItemInArray(file.columns, event.previousIndex, event.currentIndex);
    }
  }

  canExport(): boolean {
    return this.importedFiles.length > 0 &&
           this.importedFiles.some(file => 
             file.columns.some(column => column.selected)
           );
  }

  async exportAllToExcel(): Promise<void> {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Dados Combinados');
      
      // Mapa para armazenar todas as colunas únicas e seus valores
      const uniqueColumns = new Map<string, Set<string>>();
      const columnData = new Map<string, any[]>();
      
      // Primeiro passo: coletar todas as colunas únicas e seus valores
      this.importedFiles.forEach(file => {
        file.columns.forEach(column => {
          if (column.selected) {
            const normalizedName = this.normalizeColumnName(column.displayName);
            if (!uniqueColumns.has(normalizedName)) {
              uniqueColumns.set(normalizedName, new Set([column.displayName]));
              columnData.set(normalizedName, []);
            } else {
              uniqueColumns.get(normalizedName)?.add(column.displayName);
            }
          }
        });
      });
      
      // Segundo passo: coletar todos os dados para cada coluna única
      this.importedFiles.forEach(file => {
        file.data.forEach(row => {
          file.columns.forEach(column => {
            if (column.selected) {
              const normalizedName = this.normalizeColumnName(column.displayName);
              const currentData = columnData.get(normalizedName) || [];
              if (row[column.key] !== undefined && row[column.key] !== null) {
                currentData.push(row[column.key]);
              }
              columnData.set(normalizedName, currentData);
            }
          });
        });
      });
      
      // Terceiro passo: adicionar cabeçalhos e dados ao worksheet
      const headers: string[] = Array.from(uniqueColumns.keys());
      worksheet.addRow(headers);
      
      // Encontrar o maior número de linhas
      const maxRows = Math.max(...Array.from(columnData.values()).map(data => data.length));
      
      // Adicionar dados
      for (let i = 0; i < maxRows; i++) {
        const rowData = headers.map(header => {
          const data = columnData.get(header) || [];
          return data[i] || '';
        });
        worksheet.addRow(rowData);
      }
      
      // Ajustar largura das colunas
      worksheet.columns.forEach((column: any) => {
        if (column && typeof column.eachCell === 'function') {
          let maxLength = 0;
          column.eachCell({ includeEmpty: true }, (cell: any) => {
            const cellLength = cell.value ? cell.value.toString().length : 0;
            maxLength = Math.max(maxLength, cellLength);
          });
          column.width = Math.min(maxLength + 2, 50); // Limita a largura máxima a 50 caracteres
        }
      });

      // Estilizar cabeçalhos
      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6E6' }
      };

      // Gerar arquivo Excel
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, 'dados_combinados.xlsx');
    } catch (error) {
      console.error('Erro ao exportar para Excel:', error);
      this.showErrorMessage('Não foi possível exportar os dados para Excel. Tente novamente.');
    }
  }

  private normalizeColumnName(name: string): string {
    return name.toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  }

  scrollToTop(): void {
    this.topElement.nativeElement.scrollIntoView({ behavior: 'smooth' });
  }

  onDragOver(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    const dropZone = event.target as HTMLElement;
    dropZone.closest('.drop-zone')?.classList.add('dragover');
  }

  onDragLeave(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    const dropZone = event.target as HTMLElement;
    dropZone.closest('.drop-zone')?.classList.remove('dragover');
  }

  onFileDrop(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    const dropZone = event.target as HTMLElement;
    dropZone.closest('.drop-zone')?.classList.remove('dragover');

    const files = Array.from(event.dataTransfer?.files || []);
    const xmlFiles = files.filter(file => file.name.toLowerCase().endsWith('.xml'));
    
    if (xmlFiles.length > 0) {
      this.handleFiles(xmlFiles);
    }
  }

  private handleFiles(files: File[]) {
    for (const file of files) {
      const reader = new FileReader();
      reader.onload = (e: ProgressEvent<FileReader>) => {
        const xmlContent = e.target?.result as string;
        this.processXmlFile(file.name, xmlContent);
      };
      reader.readAsText(file);
    }
  }

  private async processXmlFile(fileName: string, xmlContent: string) {
    try {
      const result = await this.xmlProcessor.parseXml(xmlContent);
      if (result.length > 0) {
        const columns: Column[] = Object.keys(result[0]).map(key => ({
          key,
          displayName: key,
          selected: true
        }));

        this.importedFiles.push({
          name: fileName,
          data: result,
          columns,
          columnSearchText: ''
        });
      }
    } catch (error) {
      console.error('Erro ao processar arquivo XML:', error);
      this.showErrorMessage('Não foi possível processar o arquivo XML. Verifique se o formato está correto.');
    }
  }

  private showErrorMessage(message: string) {
    this.snackBar.open(message, 'Fechar', {
      duration: 5000,
      horizontalPosition: 'center',
      verticalPosition: 'bottom',
      panelClass: ['error-snackbar']
    });
  }

  deselectAllColumns(file: ImportedFile): void {
    file.columns.forEach(column => {
      column.selected = false;
      // Atualiza o mesmo estado em todas as instâncias da coluna em outros arquivos
      const normalizedKey = this.normalizeColumnName(column.key);
      this.importedFiles.forEach(otherFile => {
        otherFile.columns
          .filter(col => this.normalizeColumnName(col.key) === normalizedKey)
          .forEach(col => col.selected = false);
      });
    });
  }
}
