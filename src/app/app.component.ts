import { Component } from '@angular/core';
import { XmlImportComponent } from './components/xml-import/xml-import.component';

@Component({
  selector: 'app-root',
  template: '<app-xml-import></app-xml-import>',
  standalone: true,
  imports: [XmlImportComponent]
})
export class AppComponent {
  title = 'xml-import-app';
}
