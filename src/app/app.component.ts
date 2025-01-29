import { Component } from '@angular/core';
import { XmlImportComponent } from './components/xml-import/xml-import.component';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
  standalone: true,
  imports: [XmlImportComponent]
})
export class AppComponent {
  title = 'xml-import-app';
}
