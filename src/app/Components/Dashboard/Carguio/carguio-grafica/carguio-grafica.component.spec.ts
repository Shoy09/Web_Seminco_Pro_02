import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CarguioGraficaComponent } from './carguio-grafica.component';

describe('CarguioGraficaComponent', () => {
  let component: CarguioGraficaComponent;
  let fixture: ComponentFixture<CarguioGraficaComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CarguioGraficaComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(CarguioGraficaComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
