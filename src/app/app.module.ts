import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { FormsModule } from '@angular/forms';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatCardModule } from '@angular/material/card';
import { MatSnackBarModule } from '@angular/material/snack-bar';
import { DragDropModule } from '@angular/cdk/drag-drop';

import { AppComponent } from './app.component';
import { XmlImportComponent } from './components/xml-import/xml-import.component';
import { XmlProcessorService } from './services/xml-processor.service';
import { ColumnSelectionService } from './services/column-selection.service';
import { ExcelExportService } from './services/excel-export.service';

@NgModule({
  declarations: [
    AppComponent,
    XmlImportComponent
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    FormsModule,
    MatButtonModule,
    MatIconModule,
    MatCheckboxModule,
    MatInputModule,
    MatFormFieldModule,
    MatCardModule,
    MatSnackBarModule,
    DragDropModule
  ],
  providers: [
    XmlProcessorService,
    ColumnSelectionService,
    ExcelExportService
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
