import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { MatSnackBar } from '@angular/material/snack-bar';

@Injectable({
  providedIn: 'root'
})
export class ExcelExportService {
  constructor(private snackBar: MatSnackBar) {}

  async exportToExcel(data: any[], selectedColumns: string[], filename: string = 'dados_exportados.xlsx'): Promise<void> {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Dados Combinados');

      // Adicionar cabeçalhos
      worksheet.addRow(selectedColumns);

      // Adicionar dados
      data.forEach(row => {
        const rowData = selectedColumns.map(col => {
          const value = row[col];
          return value !== undefined && value !== null ? value : '';
        });
        worksheet.addRow(rowData);
      });

      // Estilizar cabeçalhos
      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6E6' }
      };

      // Ajustar largura das colunas
      worksheet.columns.forEach((column: any) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell: any) => {
          const cellLength = cell.value ? cell.value.toString().length : 0;
          maxLength = Math.max(maxLength, cellLength);
        });
        column.width = Math.min(maxLength + 2, 50);
      });

      // Gerar arquivo
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      saveAs(blob, filename);

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
      throw error;
    }
  }
}
