import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { MatDialogModule, MatDialogRef, MatDialog } from '@angular/material/dialog';
import { MatDividerModule } from '@angular/material/divider';
import { MatGridListModule } from '@angular/material/grid-list';
import { SeleccionProcesoEstatosDialogComponent } from '../../Estado/seleccion-proceso-estatos-dialog/seleccion-proceso-estatos-dialog.component';
import { OpcionesDialogComponent } from '../opciones-dialog/opciones-dialog.component';

@Component({
  selector: 'app-seleccion-proceso-dialog',
  imports: [MatDialogModule, CommonModule, MatGridListModule, MatDividerModule],
  templateUrl: './seleccion-proceso-dialog.component.html',
  styleUrl: './seleccion-proceso-dialog.component.css'
})
export class SeleccionProcesoDialogComponent {
  procesos = ['PERFORACIÓN TALADROS LARGOS', 'PERFORACIÓN HORIZONTAL', 'SOSTENIMIENTO', 'SERVICIOS AUXILIARES', 'CARGUÍO', 'ACARREO'];

  constructor(
    public dialogRef: MatDialogRef<SeleccionProcesoEstatosDialogComponent>,
    private dialog: MatDialog // 🟢 Inyectamos MatDialog aquí
  ) {}

  seleccionarProceso(proceso: string) {
    this.abrirDialogo(proceso);

  }

  cerrarDialogo() {
    this.dialogRef.close();
  }
  
  abrirDialogo(proceso: string) {
    const dialogRef = this.dialog.open(OpcionesDialogComponent, {
      data: { proceso } // 🟢 Pasamos el proceso seleccionado
    });
  
    this.dialogRef.close(); // 🔴 Cerramos el diálogo actual después de abrir el nuevo
  }
  
}