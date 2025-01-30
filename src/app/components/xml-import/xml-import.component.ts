import { Component, ElementRef, ViewChild, ChangeDetectorRef, ChangeDetectionStrategy } from '@angular/core';
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
  ],
  changeDetection: ChangeDetectionStrategy.OnPush
})
export class XmlImportComponent {
  @ViewChild('top') topElement!: ElementRef;
  importedFiles: ImportedFile[] = [];
  showBackToTop = false;

  // Mapa para controlar células expandidas
  private expandedCells = new Map<string, boolean>();

  constructor(
    private xmlProcessor: XmlProcessorService,
    private snackBar: MatSnackBar,
    private changeDetectorRef: ChangeDetectorRef
  ) {
    window.addEventListener('scroll', () => {
      this.showBackToTop = window.scrollY > 300;
    });
  }

  async onFileSelected(event: any): Promise<void> {
    const files = event.target.files;
    if (files.length === 0) return;

    try {
      for (const file of files) {
        if (!file.name.endsWith('.xml')) {
          this.snackBar.open('Por favor, selecione apenas arquivos XML.', 'Fechar', {
            duration: 5000,
            panelClass: ['warning-snackbar']
          });
          continue;
        }

        const fileContent = await this.readFileContent(file);
        const processedData = await this.xmlProcessor.processXmlContent(fileContent);
        
        if (processedData && processedData.data.length > 0) {
          const importedFile: ImportedFile = {
            name: file.name,
            data: processedData.data,
            columns: Object.keys(processedData.data[0]).map(key => ({
              key,
              displayName: key,
              selected: false
            })),
            columnSearchText: ''
          };
          
          this.importedFiles.push(importedFile);
          // Força a detecção de mudanças
          this.changeDetectorRef.detectChanges();
        }
      }

      if (this.importedFiles.length > 0) {
        this.snackBar.open('Arquivos importados com sucesso!', 'Fechar', {
          duration: 3000,
          panelClass: ['success-snackbar']
        });
        
        // Força a atualização da view
        this.changeDetectorRef.detectChanges();
        
        // Scroll para o topo após um pequeno delay
        setTimeout(() => {
          this.scrollToTop();
        }, 100);
      }
    } catch (error) {
      console.error('Erro ao processar arquivos:', error);
      this.snackBar.open('Erro ao processar arquivos. Verifique se são arquivos XML válidos.', 'Fechar', {
        duration: 5000,
        panelClass: ['error-snackbar']
      });
    }

    // Limpa o input de arquivo para permitir selecionar o mesmo arquivo novamente
    event.target.value = '';
  }

  removeFile(index: number): void {
    this.importedFiles.splice(index, 1);
  }

  getVisibleColumns(file: ImportedFile): Column[] {
    if (!file || !file.columns) return [];
    
    if (!file.columnSearchText) {
      return file.columns;
    }
    
    return file.columns.filter(column => 
      column.displayName.toLowerCase().includes(file.columnSearchText.toLowerCase())
    );
  }

  getVisibleColumnKeys(file: ImportedFile): string[] {
    return this.getVisibleColumns(file).map(col => col.key);
  }

  toggleAllColumns(checked: boolean, file: ImportedFile): void {
    console.log('toggleAllColumns', checked, file.name); // Debug log
    const newValue = checked;
    file.columns.forEach(column => {
      column.selected = newValue;
    });
    // Força a detecção de mudanças
    this.changeDetectorRef.detectChanges();
  }

  areAllColumnsSelected(file: ImportedFile): boolean {
    if (!file || !file.columns || file.columns.length === 0) return false;
    return file.columns.every(column => column.selected);
  }

  areSomeColumnsSelected(file: ImportedFile): boolean {
    if (!file || !file.columns || file.columns.length === 0) return false;
    const selectedCount = file.columns.filter(column => column.selected).length;
    return selectedCount > 0 && selectedCount < file.columns.length;
  }

  toggleColumnSelection(column: Column, file: ImportedFile): void {
    column.selected = !column.selected;
    
    // Debug
    console.log('Toggle coluna:', column.displayName);
    console.log('Estado atual:', column.selected);
    
    // Força atualização da view
    this.changeDetectorRef.detectChanges();
    
    // Verifica estado do botão
    const canExport = this.canExport();
    console.log('Pode exportar:', canExport);
  }

  onColumnDrop(event: CdkDragDrop<string[]>, file: ImportedFile): void {
    if (event.previousIndex !== event.currentIndex) {
      moveItemInArray(file.columns, event.previousIndex, event.currentIndex);
    }
  }

  canExport(): boolean {
    // Verifica se há arquivos importados
    if (this.importedFiles.length === 0) {
      return false;
    }

    // Verifica se há pelo menos uma coluna selecionada em qualquer arquivo
    const hasSelectedColumns = this.importedFiles.some(file => 
      file.columns.some(col => {
        console.log(`Coluna ${col.displayName}: ${col.selected}`);
        return col.selected;
      })
    );

    console.log('Tem colunas selecionadas:', hasSelectedColumns);
    return hasSelectedColumns;
  }

  async exportAllToExcel(): Promise<void> {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Dados Combinados');
      
      // Mapa para armazenar colunas únicas e seus valores
      const uniqueColumns = new Map<string, Set<{fileIndex: number, columnKey: string}>>();
      
      // Primeiro passo: identificar colunas únicas selecionadas
      this.importedFiles.forEach((file, fileIndex) => {
        const selectedColumns = file.columns.filter(col => col.selected);
        selectedColumns.forEach(column => {
          const normalizedName = this.normalizeColumnName(column.displayName);
          if (!uniqueColumns.has(normalizedName)) {
            uniqueColumns.set(normalizedName, new Set());
          }
          uniqueColumns.get(normalizedName)?.add({fileIndex, columnKey: column.key});
        });
      });

      // Se não houver colunas selecionadas, mostrar mensagem
      if (uniqueColumns.size === 0) {
        this.snackBar.open('Selecione pelo menos uma coluna para exportar.', 'Fechar', {
          duration: 5000,
          panelClass: ['warning-snackbar']
        });
        return;
      }

      // Adicionar cabeçalhos
      const headers = Array.from(uniqueColumns.keys());
      worksheet.addRow(headers);

      // Preparar dados combinados
      const allRows: any[][] = [];
      
      // Para cada arquivo
      this.importedFiles.forEach((file, fileIndex) => {
        file.data.forEach(dataRow => {
          while (allRows.length <= fileIndex) {
            allRows.push([]);
          }
          
          // Para cada coluna única
          headers.forEach((header, colIndex) => {
            const columnSources = uniqueColumns.get(header);
            if (columnSources) {
              // Procurar valor nas colunas correspondentes
              const values = Array.from(columnSources)
                .filter(source => source.fileIndex === fileIndex)
                .map(source => dataRow[source.columnKey])
                .filter(value => value !== undefined && value !== null);
              
              if (values.length > 0) {
                allRows[fileIndex][colIndex] = values.join(', ');
              }
            }
          });
        });
      });

      // Adicionar todas as linhas ao worksheet
      allRows.flat().forEach(row => {
        if (Object.values(row).some(value => value !== undefined)) {
          worksheet.addRow(row);
        }
      });

      // Estilizar a planilha
      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6E6' }
      };

      // Ajustar largura das colunas
      worksheet.columns.forEach((column: any) => {
        if (column && typeof column.eachCell === 'function') {
          let maxLength = 0;
          column.eachCell({ includeEmpty: true }, (cell: any) => {
            const cellLength = cell.value ? cell.value.toString().length : 0;
            maxLength = Math.max(maxLength, cellLength);
          });
          column.width = Math.min(maxLength + 2, 50);
        }
      });

      // Gerar arquivo Excel
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      saveAs(blob, 'dados_exportados.xlsx');

      this.snackBar.open('Dados exportados com sucesso!', 'Fechar', {
        duration: 3000,
        panelClass: ['success-snackbar']
      });
    } catch (error) {
      console.error('Erro ao exportar para Excel:', error);
      this.snackBar.open('Erro ao exportar para Excel. Por favor, tente novamente.', 'Fechar', {
        duration: 5000,
        panelClass: ['error-snackbar']
      });
    }
  }

  private normalizeColumnName(name: string): string {
    return name.toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  }

  isLargeContent(text: string): boolean {
    if (!text) return false;
    return text.length > 100 || text.includes('\n');
  }

  getExpandedKey(fileName: string, columnKey: string, row: any): string {
    return `${fileName}-${columnKey}-${row[columnKey]}`;
  }

  isExpanded(fileName: string, columnKey: string, row: any): boolean {
    const key = this.getExpandedKey(fileName, columnKey, row);
    return this.expandedCells.get(key) || false;
  }

  toggleExpand(fileName: string, columnKey: string, row: any): void {
    const key = this.getExpandedKey(fileName, columnKey, row);
    this.expandedCells.set(key, !this.isExpanded(fileName, columnKey, row));
    this.changeDetectorRef.detectChanges();
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

  private async readFileContent(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e: ProgressEvent<FileReader>) => {
        resolve(e.target?.result as string);
      };
      reader.onerror = reject;
      reader.readAsText(file);
    });
  }
}
