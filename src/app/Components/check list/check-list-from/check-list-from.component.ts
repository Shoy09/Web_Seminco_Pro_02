import { CommonModule } from "@angular/common";
import { Component, Inject } from "@angular/core";
import { ReactiveFormsModule, FormGroup, FormBuilder, Validators } from "@angular/forms";
import { MatDialogRef, MAT_DIALOG_DATA } from "@angular/material/dialog";
import { CheckListItem } from "../../../models/checklist-item.model";
import { CheckListItemService } from "../../../services/checklist-item.service";

@Component({
  selector: 'app-check-list-from',
  imports: [ReactiveFormsModule, CommonModule],
  templateUrl: './check-list-from.component.html',
  styleUrl: './check-list-from.component.css'
})
export class CheckListFromComponent  {
  checkListForm: FormGroup;
  mensaje: string = '';

  constructor(
    private fb: FormBuilder,
    private checkListService: CheckListItemService,
    public dialogRef: MatDialogRef<CheckListFromComponent>,
    @Inject(MAT_DIALOG_DATA) public data: { proceso: string }
  )
  {
    this.checkListForm = this.fb.group({
      proceso: [{ value: this.data.proceso, disabled: true }], // 🟢 Campo de solo lectura
      categoria: ['', Validators.required],
      nombre: ['', Validators.required]
    });
  }

  onSubmit() {
    if (this.checkListForm.valid) {
      this.checkListForm.get('proceso')?.enable(); // Habilita el campo antes de obtener el valor
      const nuevoItem: CheckListItem = this.checkListForm.value;
      this.checkListForm.get('proceso')?.disable(); // Lo vuelve a deshabilitar

      this.checkListService.createCheckListItem(nuevoItem).subscribe({
        next: (response) => {
          this.mensaje = 'Checklist creado exitosamente!';
          this.checkListForm.reset();
          this.dialogRef.close(response);
        },
        error: (error) => {
          console.error('Error al crear el checklist:', error);
          this.mensaje = 'Ocurrió un error al crear el checklist.';
        }
      });
    }
  }

  onCancel(): void {
    this.dialogRef.close(); // Cierra el diálogo sin hacer nada
  }
}