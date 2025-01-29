import { ComponentFixture, TestBed } from '@angular/core/testing';

import { XmlImportComponent } from './xml-import.component';

describe('XmlImportComponent', () => {
  let component: XmlImportComponent;
  let fixture: ComponentFixture<XmlImportComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [XmlImportComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(XmlImportComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
