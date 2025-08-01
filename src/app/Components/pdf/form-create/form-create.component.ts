import { Component, OnInit } from '@angular/core';
import { FormGroup, FormBuilder, Validators, ReactiveFormsModule } from '@angular/forms';
import { PdfService } from '../../../services/pdf.service';
import { Router } from '@angular/router';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-form-create',
  standalone: true,
  imports: [ReactiveFormsModule, CommonModule],
  templateUrl: './form-create.component.html',
  styleUrl: './form-create.component.css'
})
export class FormCreateComponent implements OnInit {
  createForm!: FormGroup;
  meses: string[] = [
    'ENERO', 'FEBRERO', 'MARZO', 'ABRIL',
    'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
    'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'
  ];
  procesos: string[] = ['PERFORACIÓN TALADROS LARGOS', 'PERFORACIÓN HORIZONTAL', 'SOSTENIMIENTO', 'SERVICIOS AUXILIARES', 'CARGUÍO', 'ACARREO'];
  
  // Nuevas opciones para los campos
  tiposLabor: string[] = ['TIPO 1', 'TIPO 2', 'TIPO 3']; // Ajusta según tus necesidades
  labores: string[] = ['LABOR 1', 'LABOR 2', 'LABOR 3']; // Ajusta según tus necesidades
  alas: string[] = ['ALA A', 'ALA B', 'ALA C', 'ALA D']; // Ajusta según tus necesidades
  
  pdfFile: File | null = null;

  constructor(
    private fb: FormBuilder,
    private pdfService: PdfService,
    private router: Router
  ) {}

  ngOnInit(): void {
    const mesActual = this.meses[new Date().getMonth()];
    this.createForm = this.fb.group({
      proceso: [this.procesos[0], Validators.required],
      mes: [mesActual, Validators.required],
      tipo_labor: [this.tiposLabor[0], Validators.required],
      labor: [this.labores[0], Validators.required],
      ala: [this.alas[0], Validators.required],
      pdf: [null, Validators.required]
    });
  }

  onFileSelected(event: any): void {
    const file: File = event.target.files[0];
    if (file && file.type === 'application/pdf') {
      this.pdfFile = file;
      this.createForm.patchValue({ pdf: file });
    } else {
      alert('Por favor selecciona un archivo PDF válido.');
    }
  }

  onSubmit(): void {
    if (this.createForm.valid && this.pdfFile) {
      const formData = new FormData();
      formData.append('proceso', this.createForm.get('proceso')?.value.toUpperCase());
      formData.append('mes', this.createForm.get('mes')?.value.toUpperCase());
      formData.append('tipo_labor', this.createForm.get('tipo_labor')?.value.toUpperCase());
      formData.append('labor', this.createForm.get('labor')?.value.toUpperCase());
      formData.append('ala', this.createForm.get('ala')?.value.toUpperCase());
      formData.append('archivo', this.pdfFile);

      this.pdfService.createPdf(formData).subscribe({
        next: (res) => {
          alert('PDF subido exitosamente.');
          this.createForm.reset({
            proceso: this.procesos[0],
            mes: this.meses[new Date().getMonth()],
            tipo_labor: this.tiposLabor[0],
            labor: this.labores[0],
            ala: this.alas[0]
          });
          this.pdfFile = null;
        },
        error: (err) => {
          console.error('Error al subir PDF:', err);
          alert('Ocurrió un error al subir el PDF.');
        }
      });
    } else {
      alert('Por favor completa todos los campos y selecciona un PDF.');
    }
  }
}