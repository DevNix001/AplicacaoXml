import { Injectable } from '@angular/core';
import { BehaviorSubject } from 'rxjs';

interface ColumnSelection {
  fileIndex: number;
  columnKey: string;
  selected: boolean;
}

@Injectable({
  providedIn: 'root'
})
export class ColumnSelectionService {
  private selectedColumnsSubject = new BehaviorSubject<ColumnSelection[]>([]);
  selectedColumns$ = this.selectedColumnsSubject.asObservable();

  constructor() {
    console.log('ColumnSelectionService initialized');
  }

  toggleColumnSelection(fileIndex: number, columnKey: string, selected: boolean): void {
    console.log('Toggling column selection:', { fileIndex, columnKey, selected });
    const currentSelections = this.selectedColumnsSubject.value;
    const existingIndex = currentSelections.findIndex(
      sel => sel.fileIndex === fileIndex && sel.columnKey === columnKey
    );

    if (existingIndex >= 0) {
      if (!selected) {
        currentSelections.splice(existingIndex, 1);
      } else {
        currentSelections[existingIndex].selected = selected;
      }
    } else if (selected) {
      currentSelections.push({ fileIndex, columnKey, selected });
    }

    console.log('Updated selections:', currentSelections);
    this.selectedColumnsSubject.next([...currentSelections]);
  }

  hasSelectedColumns(): boolean {
    const hasSelected = this.selectedColumnsSubject.value.some(sel => sel.selected);
    console.log('Has selected columns:', hasSelected);
    return hasSelected;
  }

  getSelectedColumns(): string[] {
    const columns = this.selectedColumnsSubject.value
      .filter(sel => sel.selected)
      .map(sel => sel.columnKey);
    console.log('Getting selected columns:', columns);
    return columns;
  }

  clearSelections(): void {
    console.log('Clearing all selections');
    this.selectedColumnsSubject.next([]);
  }
}
