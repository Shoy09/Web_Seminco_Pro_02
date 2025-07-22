import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SeleccionProcesoDialogComponent } from './seleccion-proceso-dialog.component';

describe('SeleccionProcesoDialogComponent', () => {
  let component: SeleccionProcesoDialogComponent;
  let fixture: ComponentFixture<SeleccionProcesoDialogComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SeleccionProcesoDialogComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SeleccionProcesoDialogComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
