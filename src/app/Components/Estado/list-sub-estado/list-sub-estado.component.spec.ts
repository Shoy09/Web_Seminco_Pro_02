import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ListSubEstadoComponent } from './list-sub-estado.component';

describe('ListSubEstadoComponent', () => {
  let component: ListSubEstadoComponent;
  let fixture: ComponentFixture<ListSubEstadoComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ListSubEstadoComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ListSubEstadoComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
