import { Component, Inject } from '@angular/core';
import { MAT_DIALOG_DATA, MatDialog, MatDialogModule, MatDialogRef } from '@angular/material/dialog';
import { MatIconModule } from '@angular/material/icon';
import * as XLSX from 'xlsx';
import { LoadingDialogComponent } from '../../Reutilizables/loading-dialog/loading-dialog.component';
import { CheckListFromComponent } from '../check-list-from/check-list-from.component';
import { CheckListItemService } from '../../../services/checklist-item.service';
import { CheckListItem } from '../../../models/checklist-item.model';

@Component({
  selector: 'app-opciones-dialog',
  standalone: true,
  imports: [MatDialogModule, MatIconModule],
  templateUrl: './opciones-dialog.component.html',
  styleUrl: './opciones-dialog.component.css'
})
export class OpcionesDialogComponent {
  constructor(
    private checkdService: CheckListItemService,
    public dialogRef: MatDialogRef<OpcionesDialogComponent>,
    private dialog: MatDialog,
    @Inject(MAT_DIALOG_DATA) public data: { proceso: string }
  ) {}

  seleccionar(opcion: string) {
    if (opcion === 'estado') {
      this.dialogRef.close();
      this.dialog.open(CheckListFromComponent, {
        width: '400px',
        data: { proceso: this.data.proceso }
      });
    } else if (opcion === 'excel') {
      this.abrirExploradorArchivos();
    } else {
      this.dialogRef.close(opcion);
    }
  }

  abrirExploradorArchivos() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xls, .xlsx';
    input.style.display = 'none';

    input.addEventListener('change', async (event: any) => {
      const archivo = event.target.files[0];
      if (archivo) {
        this.procesarArchivoExcel(archivo);
        this.dialogRef.close();
      }
    });

    document.body.appendChild(input);
    input.click();
    document.body.removeChild(input);
  }

async procesarArchivoExcel(archivo: File) {
  const reader = new FileReader();

  reader.onload = async (e: any) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    if (!worksheet) {
      console.error(`No se encontró la hoja llamada "${firstSheetName}"`);
      return;
    }

    const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Omitir encabezados y filtrar filas válidas
    const checklistItems: CheckListItem[] = jsonData
      .slice(1) // Omitir encabezado
      .filter(fila => fila[0] && fila[1]) // Solo filas con categoria y nombre
      .map((fila) => ({
        proceso: this.data.proceso,
        categoria: fila[0].toString().trim(), // Asegurar string y limpiar espacios
        nombre: fila[1].toString().trim()
      }));

    if (checklistItems.length === 0) {
      console.warn('No hay items válidos para procesar.');
      return;
    }

    console.log(`Procesando ${checklistItems.length} items válidos...`);
    
    this.mostrarPantallaCarga();
    await this.enviarItemsALaBD(checklistItems);
    this.cerrarPantallaCarga();
  };

  reader.readAsArrayBuffer(archivo);
}

  async enviarItemsALaBD(items: CheckListItem[]) {
    let itemsInsertados = 0;

    for (const item of items) {
      try {
        await this.checkdService.createCheckListItem(item).toPromise();
        itemsInsertados++;
      } catch (error) {
        console.error('Error al insertar item:', item.nombre, error);
      }
    }

    console.log(`${itemsInsertados} items insertados correctamente.`);
  }

  mostrarPantallaCarga() {
    this.dialog.open(LoadingDialogComponent, {
      disableClose: true
    });
  }

  cerrarPantallaCarga() {
    this.dialog.closeAll();
  }
}
