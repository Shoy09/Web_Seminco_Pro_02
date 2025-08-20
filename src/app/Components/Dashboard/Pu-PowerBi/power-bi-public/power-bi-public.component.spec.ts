import { ComponentFixture, TestBed } from '@angular/core/testing';

import { PowerBiPublicComponent } from './power-bi-public.component';

describe('PowerBiPublicComponent', () => {
  let component: PowerBiPublicComponent;
  let fixture: ComponentFixture<PowerBiPublicComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [PowerBiPublicComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(PowerBiPublicComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
