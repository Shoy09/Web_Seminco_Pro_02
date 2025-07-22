import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CheckListListaComponent } from './check-list-lista.component';

describe('CheckListListaComponent', () => {
  let component: CheckListListaComponent;
  let fixture: ComponentFixture<CheckListListaComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CheckListListaComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(CheckListListaComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
