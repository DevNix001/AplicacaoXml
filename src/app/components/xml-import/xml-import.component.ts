import { Component, ChangeDetectorRef, ViewChild, ElementRef, OnInit, OnDestroy, HostListener } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatCardModule } from '@angular/material/card';
import { MatSnackBar } from '@angular/material/snack-bar';
import { DragDropModule, CdkDragDrop, moveItemInArray } from '@angular/cdk/drag-drop';
import { XmlProcessorService } from '../../services/xml-processor.service';
import { ColumnSelectionService } from '../../services/column-selection.service';
import { ExcelExportService } from '../../services/excel-export.service';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Subscription } from 'rxjs';

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
  index: number;
}

@Component({
  selector: 'app-xml-import',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatButtonModule,
    MatIconModule,
    MatCheckboxModule,
    MatInputModule,
    MatFormFieldModule,
    MatCardModule,
    DragDropModule
  ],
  templateUrl: './xml-import.component.html',
  styleUrls: ['./xml-import.component.scss']
})
export class XmlImportComponent implements OnInit, OnDestroy {
  @ViewChild('top') topElement!: ElementRef;
  importedFiles: ImportedFile[] = [];
  private subscriptions: Subscription = new Subscription();
  private expandedCells = new Map<string, boolean>();
  showBackToTop = false;

  constructor(
    private xmlProcessor: XmlProcessorService,
    public columnSelectionService: ColumnSelectionService,
    private excelExportService: ExcelExportService,
    private snackBar: MatSnackBar,
    private changeDetectorRef: ChangeDetectorRef
  ) {}

  @HostListener('window:scroll', [])
  onWindowScroll() {
    this.showBackToTop = window.scrollY > 300;
  }

  ngOnInit() {
    this.subscriptions.add(
      this.columnSelectionService.selectedColumns$.subscribe(() => {
        this.changeDetectorRef.detectChanges();
      })
    );
  }

  ngOnDestroy() {
    this.subscriptions.unsubscribe();
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

        const processedData = await this.xmlProcessor.processXmlFile(file);
        
        if (processedData && processedData.data.length > 0) {
          const importedFile: ImportedFile = {
            name: file.name,
            data: processedData.data,
            columns: Object.keys(processedData.data[0]).map(key => ({
              key,
              displayName: key,
              selected: false
            })),
            columnSearchText: '',
            index: this.importedFiles.length
          };
          
          this.importedFiles.push(importedFile);
        }
      }

      if (this.importedFiles.length > 0) {
        this.snackBar.open('Arquivos importados com sucesso!', 'Fechar', {
          duration: 3000,
          panelClass: ['success-snackbar']
        });
        
        this.changeDetectorRef.detectChanges();
        setTimeout(() => this.scrollToTop(), 100);
      }
    } catch (error) {
      console.error('Erro ao processar arquivos:', error);
      this.snackBar.open('Erro ao processar arquivos. Verifique se são arquivos XML válidos.', 'Fechar', {
        duration: 5000,
        panelClass: ['error-snackbar']
      });
    }

    event.target.value = '';
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    const dropZone = event.target as HTMLElement;
    dropZone.closest('.drop-zone')?.classList.add('dragover');
  }

  onDragLeave(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    const dropZone = event.target as HTMLElement;
    dropZone.closest('.drop-zone')?.classList.remove('dragover');
  }

  onFileDrop(event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    const dropZone = event.target as HTMLElement;
    dropZone.closest('.drop-zone')?.classList.remove('dragover');

    const files = Array.from(event.dataTransfer?.files || []);
    if (files.length > 0) {
      this.handleFiles(files);
    }
  }

  private async handleFiles(files: File[]): Promise<void> {
    const xmlFiles = files.filter(file => file.name.toLowerCase().endsWith('.xml'));
    
    if (xmlFiles.length === 0) {
      this.snackBar.open('Por favor, selecione apenas arquivos XML.', 'Fechar', {
        duration: 5000,
        panelClass: ['warning-snackbar']
      });
      return;
    }

    try {
      for (const file of xmlFiles) {
        const processedData = await this.xmlProcessor.processXmlFile(file);
        if (processedData && processedData.data.length > 0) {
          this.importedFiles.push({
            name: file.name,
            data: processedData.data,
            columns: Object.keys(processedData.data[0]).map(key => ({
              key,
              displayName: key,
              selected: false
            })),
            columnSearchText: '',
            index: this.importedFiles.length
          });
        }
      }

      if (this.importedFiles.length > 0) {
        this.snackBar.open('Arquivos importados com sucesso!', 'Fechar', {
          duration: 3000,
          panelClass: ['success-snackbar']
        });
      }
    } catch (error) {
      console.error('Erro ao processar arquivos:', error);
      this.snackBar.open('Erro ao processar arquivos. Verifique se são arquivos XML válidos.', 'Fechar', {
        duration: 5000,
        panelClass: ['error-snackbar']
      });
    }
  }

  getVisibleColumns(file: ImportedFile): Column[] {
    if (!file.columnSearchText) {
      return file.columns;
    }
    const searchText = file.columnSearchText.toLowerCase();
    return file.columns.filter(col => 
      col.displayName.toLowerCase().includes(searchText)
    );
  }

  toggleColumnSelection(column: Column, file: ImportedFile): void {
    column.selected = !column.selected;
    this.columnSelectionService.toggleColumnSelection(
      file.index,
      column.key,
      column.selected
    );
  }

  toggleAllColumns(file: ImportedFile): void {
    const allSelected = this.areAllColumnsSelected(file);
    file.columns.forEach(column => {
      column.selected = !allSelected;
      this.columnSelectionService.toggleColumnSelection(file.index, column.key, !allSelected);
    });
  }

  areAllColumnsSelected(file: ImportedFile): boolean {
    return file.columns.length > 0 && file.columns.every(column => column.selected);
  }

  areSomeColumnsSelected(file: ImportedFile): boolean {
    const selectedCount = file.columns.filter(column => column.selected).length;
    return selectedCount > 0 && selectedCount < file.columns.length;
  }

  onColumnDrop(event: CdkDragDrop<string[]>, file: ImportedFile): void {
    moveItemInArray(file.columns, event.previousIndex, event.currentIndex);
    this.changeDetectorRef.detectChanges();
  }

  isLargeContent(text: string): boolean {
    if (!text) return false;
    return text.length > 100 || text.includes('\n');
  }

  isExpanded(fileName: string, columnKey: string, row: any): boolean {
    const key = `${fileName}-${columnKey}-${JSON.stringify(row)}`;
    return this.expandedCells.get(key) || false;
  }

  toggleExpand(fileName: string, columnKey: string, row: any): void {
    const key = `${fileName}-${columnKey}-${JSON.stringify(row)}`;
    this.expandedCells.set(key, !this.isExpanded(fileName, columnKey, row));
    this.changeDetectorRef.detectChanges();
  }

  scrollToTop(): void {
    this.topElement?.nativeElement?.scrollIntoView({ behavior: 'smooth' });
  }

  private normalizeColumnName(name: string): string {
    return name.toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  }

  async exportAllToExcel(): Promise<void> {
    try {
      // Verificar seleções
      const selectedColumns = new Set<string>();
      this.importedFiles.forEach(file => {
        console.log('Checking file:', file.name);
        file.columns.forEach(col => {
          console.log('Column:', col.key, 'Selected:', col.selected);
          if (col.selected) {
            selectedColumns.add(col.key);
          }
        });
      });

      console.log('Selected columns:', Array.from(selectedColumns));

      if (selectedColumns.size === 0) {
        this.snackBar.open('Selecione pelo menos uma coluna para exportar.', 'Fechar', {
          duration: 5000,
          panelClass: ['warning-snackbar']
        });
        return;
      }

      // Combinar dados
      const allData: any[] = [];
      this.importedFiles.forEach(file => {
        file.data.forEach(row => {
          const processedRow: any = {};
          Array.from(selectedColumns).forEach(colKey => {
            processedRow[colKey] = row[colKey] || '';
          });
          allData.push(processedRow);
        });
      });

      console.log('Data to export:', allData);

      await this.excelExportService.exportToExcel(
        allData,
        Array.from(selectedColumns),
        'dados_exportados.xlsx'
      );
    } catch (error) {
      console.error('Erro ao exportar:', error);
      this.snackBar.open('Erro ao exportar dados. Por favor, tente novamente.', 'Fechar', {
        duration: 5000,
        panelClass: ['error-snackbar']
      });
    }
  }

  removeFile(index: number): void {
    this.importedFiles.splice(index, 1);
    this.changeDetectorRef.detectChanges();
  }
}
