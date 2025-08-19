import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CreateSubEstadoComponent } from './create-sub-estado.component';

describe('CreateSubEstadoComponent', () => {
  let component: CreateSubEstadoComponent;
  let fixture: ComponentFixture<CreateSubEstadoComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CreateSubEstadoComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(CreateSubEstadoComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
