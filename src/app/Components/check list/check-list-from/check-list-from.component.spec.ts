import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CheckListFromComponent } from './check-list-from.component';

describe('CheckListFromComponent', () => {
  let component: CheckListFromComponent;
  let fixture: ComponentFixture<CheckListFromComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CheckListFromComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(CheckListFromComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
