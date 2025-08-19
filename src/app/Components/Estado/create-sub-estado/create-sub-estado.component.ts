import { Component, Inject, OnInit } from '@angular/core';
import { MAT_DIALOG_DATA, MatDialogRef } from '@angular/material/dialog';
import { FormBuilder, FormGroup, Validators, ReactiveFormsModule } from '@angular/forms';
import { EstadoService } from '../../../services/estado.service';
import { CommonModule } from '@angular/common';
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatButtonModule } from '@angular/material/button';
import { MatDialogModule } from '@angular/material/dialog';
import { SubEstado } from '../../../models/Estado';

@Component({
  selector: 'app-create-sub-estado',
  standalone: true,
  imports: [
    CommonModule,
    ReactiveFormsModule,
    MatInputModule,
    MatFormFieldModule,
    MatButtonModule,
    MatDialogModule
  ],
  templateUrl: './create-sub-estado.component.html',
  styleUrls: ['./create-sub-estado.component.css']
})
export class CreateSubEstadoComponent implements OnInit {
  subEstadoForm: FormGroup;
  loading = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: { idEstado: number },
    private dialogRef: MatDialogRef<CreateSubEstadoComponent>,
    private estadoService: EstadoService,
    private fb: FormBuilder
  ) {
    this.subEstadoForm = this.fb.group({
      codigo: ['', [Validators.required, Validators.maxLength(50)]],
      tipo_estado: ['', [Validators.required, Validators.maxLength(100)]]
    });
  }

  ngOnInit(): void {}

  onSubmit(): void {
    if (this.subEstadoForm.valid) {
      this.loading = true;
      const subEstadoData: SubEstado = {
        id: 0, // El backend asignará el ID
        codigo: this.subEstadoForm.value.codigo,
        tipo_estado: this.subEstadoForm.value.tipo_estado,
        estadoId: this.data.idEstado
      };

      this.estadoService.createSubEstado(this.data.idEstado, subEstadoData).subscribe({
        next: (response) => {
          this.dialogRef.close(true); // Cierra el diálogo y devuelve true indicando éxito
        },
        error: (error) => {
          console.error('Error al crear subestado:', error);
          this.loading = false;
        }
      });
    }
  }

  onCancel(): void {
    this.dialogRef.close(false); // Cierra el diálogo sin acción
  }
}