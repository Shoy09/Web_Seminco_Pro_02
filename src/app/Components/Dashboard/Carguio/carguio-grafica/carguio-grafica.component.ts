import { Component, OnInit } from '@angular/core';
import { NubeOperacion } from '../../../../models/operaciones.models';
import { Meta } from '@angular/platform-browser';
import { ToastrService } from 'ngx-toastr';
import { MetaService } from '../../../../services/meta.service';
import { OperacionService } from '../../../../services/OperacionService .service';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelCarguioExportService } from '../../../../services/export/general/ExceCarguiolExportService.service';
import { ExcelCarguioExportServiceFiltro } from '../../../../services/export/filtro/ExceCarguiolExportService.service';
import { ExcelCarguioImportService } from '../../../../services/import/excel-import-carguio.service';

@Component({
  selector: 'app-carguio-grafica',
  imports: [FormsModule, CommonModule],
  templateUrl: './carguio-grafica.component.html',
  styleUrl: './carguio-grafica.component.css'
})
export class CarguioGraficaComponent implements OnInit {
  datosOperaciones: NubeOperacion[] = [];
  datosOperacionesExport: NubeOperacion[] = [];
  datosOperacionesOriginal: NubeOperacion[] = [];

  fechaDesde: string = '';
fechaHasta: string = '';
turnoSeleccionado: string = '';
turnos: string[] = ['DÍA', 'NOCHE'];

  constructor(private _toastr: ToastrService, private metaService: MetaService, private operacionService: OperacionService,private excelHorizontalExportService: ExcelCarguioExportService, private excelHorizontalExportServicefiltro: ExcelCarguioExportServiceFiltro, private excelImport: ExcelCarguioImportService) {}

  ngOnInit(): void {
    const fechaISO = this.obtenerFechaLocalISO();
    this.fechaDesde = fechaISO;
    this.fechaHasta = fechaISO;
    this.turnoSeleccionado = this.obtenerTurnoActual();

    this.obtenerDatos();
  }

  obtenerDatos(): void {
    this.operacionService.getOperacionesCargui().subscribe({
      next: (data) => {
        this.datosOperacionesOriginal = data;
        this.datosOperacionesExport = data;

        // Aplicar filtros por fecha actual y turno automáticamente
        const filtros = {
          fechaDesde: this.fechaDesde,
          fechaHasta: this.fechaHasta,
          turnoSeleccionado: this.turnoSeleccionado
        };

        this.datosOperaciones = this.filtrarDatos(this.datosOperacionesOriginal, filtros);

        this._toastr.success('Datos cargados correctamente Carguio', 'Éxito');

      },
      error: (err) => {
        this._toastr.error('No se pudieron obtener los datos', 'Error');
        console.error('❌ Error al obtener datos:', err);
      }
    });
  }

private obtenerMesDeFecha(fecha: string): string {
  if (!fecha) return '';

  const meses = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];

  // Dividir la fecha y crear el Date objeto en UTC para evitar problemas de zona horaria
  const partes = fecha.split('-');
  const year = parseInt(partes[0], 10);
  const month = parseInt(partes[1], 10) - 1; // Restamos 1 porque los meses en Date son 0-based
  const day = parseInt(partes[2], 10);

  // Crear la fecha en UTC
  const date = new Date(Date.UTC(year, month, day));

  return meses[date.getUTCMonth()]; // Usamos getUTCMonth() para obtener el mes correcto
}


  obtenerTurnoActual(): string {
    const ahora = new Date();
    const hora = ahora.getHours();

    // Turno de día: 7:00 AM a 6:59 PM (07:00 - 18:59)
    if (hora >= 7 && hora < 19) {
      return 'DÍA';
    } else {
      // Turno de noche: 7:00 PM a 6:59 AM
      return 'NOCHE';
    }
  }

  quitarFiltros(): void {
    const fechaISO = this.obtenerFechaLocalISO();
    this.fechaDesde = fechaISO;
    this.fechaHasta = fechaISO;
    this.turnoSeleccionado = this.obtenerTurnoActual();

    const filtros = {
      fechaDesde: this.fechaDesde,
      fechaHasta: this.fechaHasta,
      turnoSeleccionado: this.turnoSeleccionado
    };

    this.datosOperaciones = this.filtrarDatos(this.datosOperacionesOriginal, filtros);
  }

  obtenerFechaLocalISO(): string {
    const hoy = new Date();
    const año = hoy.getFullYear();
    const mes = (hoy.getMonth() + 1).toString().padStart(2, '0'); // meses comienzan en 0
    const dia = hoy.getDate().toString().padStart(2, '0');
    return `${año}-${mes}-${dia}`;
  }


  aplicarFiltrosLocales(): void {
    // Crear objeto con los filtros actuales
    const filtros = {
      fechaDesde: this.fechaDesde,
      fechaHasta: this.fechaHasta,
      turnoSeleccionado: this.turnoSeleccionado
    };

    // Aplicar filtros a los datos ORIGINALES (this.datosOperacionesOriginal)
    const datosFiltrados = this.filtrarDatos(this.datosOperacionesOriginal, filtros);

    // Actualizar los datos filtrados
    this.datosOperaciones = datosFiltrados;
  }

private obtenerCantidadDias(fechaInicio: string, fechaFin: string): number {
  const inicio = new Date(fechaInicio);
  const fin = new Date(fechaFin);
  const diffTime = Math.abs(fin.getTime() - inicio.getTime());
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 para incluir ambos días
  return diffDays;
}


  filtrarDatos(datos: NubeOperacion[], filtros: any): NubeOperacion[] {
    return datos.filter(operacion => {
      const fechaOperacion = new Date(operacion.fecha);
      const fechaDesde = filtros.fechaDesde ? new Date(filtros.fechaDesde) : null;
      const fechaHasta = filtros.fechaHasta ? new Date(filtros.fechaHasta) : null;

      // Verificar si la fecha de operación está dentro del rango
      if (fechaDesde && fechaOperacion < fechaDesde) {
        return false;
      }

      if (fechaHasta && fechaOperacion > fechaHasta) {
        return false;
      }

      // Verificar si el turno de la operación coincide con el turno seleccionado
      if (filtros.turnoSeleccionado && operacion.turno !== filtros.turnoSeleccionado) {
        return false;
      }

      return true;
    });
  }


  

exportToExcelHorizontal() {
  this.excelHorizontalExportService.exportOperacionesToExcel(
    this.datosOperacionesExport,
    'Reporte_Operaciones'
  );
}

exportToExcelHorizontalfiltro() {
  this.excelHorizontalExportServicefiltro.exportOperacionesToExcel(
    this.datosOperaciones,
    'Reporte_Operaciones'
  );
}

onFileSelected(event: any) {
  const file: File = event.target.files[0];
  if (!file) return;

  this.excelImport.importOperacionesFromExcel(file).then(operaciones => {
    // console.log('Operaciones importadas:', operaciones);

    // 🔹 Enviar el JSON directamente al POST
    this.operacionService.postOperacionesCargui(operaciones).subscribe({
      next: response => {
        console.log('Operaciones enviadas correctamente:', response);
        this._toastr.success('Operaciones enviadas a la API correctamente.', 'Éxito');
      },
      error: err => {
        console.error('Error al enviar operaciones:', err);
        this._toastr.error('Ocurrió un error al enviar operaciones.', 'Error');
      }
    });

  });
}

}
