import { Component, Inject, OnInit } from '@angular/core';
import { MAT_DIALOG_DATA, MatDialog, MatDialogModule, MatDialogRef } from '@angular/material/dialog';
import { MatTableModule } from '@angular/material/table';
import { CommonModule } from '@angular/common';
import { EstadoService } from '../../../services/estado.service';
import { ConfirmDialogComponent } from '../../Reutilizables/confirm-dialog/confirm-dialog.component';
import { SubEstado } from '../../../models/Estado';

@Component({
  selector: 'app-list-sub-estado',
  imports: [CommonModule, MatDialogModule, MatTableModule],
  templateUrl: './list-sub-estado.component.html',
  styleUrl: './list-sub-estado.component.css'
})
export class ListSubEstadoComponent implements OnInit {
  displayedColumns: string[] = ['codigo', 'tipo_estado', 'acciones'];
  dataSource: SubEstado[] = [];
  loading = true;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: { idEstado: number },
    private estadoService: EstadoService,
    private dialog: MatDialog,
    public dialogRef: MatDialogRef<ListSubEstadoComponent> // Cambiado de private a public
  ) {}

  ngOnInit(): void {
    this.loadSubEstados();
  }

  loadSubEstados(): void {
    this.estadoService.getSubEstadosByEstadoId(this.data.idEstado).subscribe({
      next: (subEstados) => {
        this.dataSource = subEstados;
        this.loading = false;
      },
      error: (error) => {
        console.error('Error al cargar subestados:', error);
        this.loading = false;
      }
    });
  }

 eliminarSubEstado(id: number): void {
    const confirmDialogRef = this.dialog.open(ConfirmDialogComponent, {
      width: '350px',
      data: { 
        mensaje: '¿Estás seguro de que deseas eliminar este subestado?',
        titulo: 'Confirmar eliminación'
      }
    });

    confirmDialogRef.afterClosed().subscribe((confirmado: boolean) => {
      if (confirmado) {
        this.loading = true;
        this.estadoService.deleteSubEstado(id).subscribe({
          next: () => {
            this.loadSubEstados(); // Recargar la lista después de eliminar
          },
          error: (error) => {
            console.error('Error al eliminar subestado:', error);
            this.loading = false;
          }
        });
      }
    });
  }
}