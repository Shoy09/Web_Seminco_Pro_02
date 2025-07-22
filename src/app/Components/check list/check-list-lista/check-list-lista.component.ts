import { Component, OnInit, ViewChild } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { MatPaginator, MatPaginatorModule } from '@angular/material/paginator';
import { MatTableDataSource, MatTableModule } from '@angular/material/table';
import { CheckListItem } from '../../../models/checklist-item.model';
import { CheckListItemService } from '../../../services/checklist-item.service';
import { ReactiveFormsModule } from '@angular/forms';
import { ConfirmDialogComponent } from '../../Reutilizables/confirm-dialog/confirm-dialog.component';
import { SeleccionProcesoDialogComponent } from '../seleccion-proceso-dialog/seleccion-proceso-dialog.component';

@Component({
  selector: 'app-check-list-lista',
  imports: [ReactiveFormsModule, MatTableModule, MatPaginatorModule],
  templateUrl: './check-list-lista.component.html',
  styleUrl: './check-list-lista.component.css'
})
export class CheckListListaComponent implements OnInit {
   displayedColumns: string[] = ['proceso', 'categoria', 'nombre', 'acciones'];
  dataSource = new MatTableDataSource<CheckListItem>();

  @ViewChild(MatPaginator) paginator!: MatPaginator;

  constructor(private checkService: CheckListItemService, public dialog: MatDialog) {}

  ngOnInit(): void {
    this.getEstados();
    this.checkService.getItemsActualizados().subscribe(cambio => {
      if (cambio) {
        this.getEstados(); // Recargar la tabla cuando haya cambios
      }
    });
  }
  

  getEstados(): void {
    this.checkService.getCheckListItems().subscribe(
      (data: CheckListItem[]) => {
        this.dataSource.data = data;
  
        // Importante: Forzar actualización del paginador
        if (this.paginator) {
          this.dataSource.paginator = this.paginator;
          this.paginator.firstPage(); // Regresar a la primera página para evitar inconsistencias
        }
      },
      (error: any) => {
        console.error('Error al obtener los estados', error);
      }
    );
  }
  

  applyFilter(event: Event) {
    const filterValue = (event.target as HTMLInputElement).value.trim().toLowerCase();
    this.dataSource.filter = filterValue;
  }

  
  abrirDialogoEditar(estado: CheckListItem) {
    // const dialogRef = this.dialog.open(EstadoFormEditarComponent, {
    //   width: '700px',
    //   data: estado, // Pasamos el estado seleccionado
    //   autoFocus: false
    // });
  
    // dialogRef.afterClosed().subscribe((estadoEditado) => {
    //   if (estadoEditado) {
    //     this.getEstados(); // Volver a cargar los estados si hubo cambios
    //   }
    // });
  }
  
  eliminarItem(id: number): void {
    const dialogRef = this.dialog.open(ConfirmDialogComponent, {
      width: '350px',
      data: { mensaje: '¿Estás seguro de que deseas eliminar?' }
    });

    dialogRef.afterClosed().subscribe((confirmado: boolean) => {
      if (confirmado) {
        this.checkService.deleteCheckListItem(id).subscribe(
          () => {
            this.getEstados(); // Refrescar la lista después de eliminar
          },
          (error) => {
            console.error('Error al eliminar el estado', error);
          }
        );
      }
    });
  }

  abrirSeleccionProcesoDialogo() {
    const dialogRef = this.dialog.open(SeleccionProcesoDialogComponent, {
      width: '400px'
    });
  
    dialogRef.afterClosed().subscribe((procesoSeleccionado) => {
      if (procesoSeleccionado) {
        
      }
    });
  }
  

  // abrirDialogoOpciones() {
  //   const dialogRef = this.dialog.open(OpcionesDialogComponent, {
  //     width: '400px'
  //   });
  
  //   dialogRef.afterClosed().subscribe((opcion) => {
  //     if (opcion === 'estado') {
  //       this.abrirDialogo();
  //     } else if (opcion === 'excel') {
  //       this.abrirExploradorArchivos();
  //     }
  //   });
  // }
  
}