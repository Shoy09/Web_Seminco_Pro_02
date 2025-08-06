import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeOperacion } from '../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelHorizontalExportService {
  exportOperacionesToExcel(datosOperacionesExport: NubeOperacion[], fileName: string) {
    // Preparar datos para cada hoja
    const ejecutadoData = this.prepareEjecutadoData(datosOperacionesExport);
    const estadosData = this.prepareEstadosData(datosOperacionesExport);
    const checklistData = this.prepareChecklistData(datosOperacionesExport);

    // Crear un nuevo libro de trabajo
    const wb: XLSX.WorkBook = XLSX.utils.book_new();

    // Añadir hojas al libro
    const ejecutadoWS = XLSX.utils.json_to_sheet(ejecutadoData);
    const estadosWS = XLSX.utils.json_to_sheet(estadosData);
    const checklistWS = XLSX.utils.json_to_sheet(checklistData);

    // Ajustar el ancho de las columnas
    this.adjustColumnWidth(ejecutadoWS, ejecutadoData);
    this.adjustColumnWidth(estadosWS, estadosData);
    this.adjustColumnWidth(checklistWS, checklistData);

    XLSX.utils.book_append_sheet(wb, ejecutadoWS, 'Ejecutado');
    XLSX.utils.book_append_sheet(wb, estadosWS, 'Estados');
    XLSX.utils.book_append_sheet(wb, checklistWS, 'Checklist');

    // Exportar el archivo
    XLSX.writeFile(wb, `${fileName}_Horizontal_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  private prepareEjecutadoData(operaciones: NubeOperacion[]): any[] {
    const data: any[] = [];
    
    operaciones.forEach(op => {
      if (op.horometros && op.horometros.length > 0) {
        op.horometros.forEach(horometro => {
          data.push({
            'ID Operación': op.id,
            'Turno': op.turno,
            'Equipo': op.equipo,
            'Código': op.codigo,
            'Empresa': op.empresa,
            'Fecha': op.fecha,
            'Tipo Operación': op.tipo_operacion,
            'Estado': op.estado,
            'Envío': op.envio,
            'Horómetro - Nombre': horometro.nombre,
            'Horómetro - Inicial': horometro.inicial,
            'Horómetro - Final': horometro.final,
            'Horómetro - OP': horometro.EstaOP ? 'Sí' : 'No',
            'Horómetro - INOP': horometro.EstaINOP ? 'Sí' : 'No',
            'Horómetro - ID': horometro.id
          });
        });
      } else {
        data.push({
          'ID Operación': op.id,
          'Turno': op.turno,
          'Equipo': op.equipo,
          'Código': op.codigo,
          'Empresa': op.empresa,
          'Fecha': op.fecha,
          'Tipo Operación': op.tipo_operacion,
          'Estado': op.estado,
          'Envío': op.envio,
          'Horómetro - Nombre': 'N/A',
          'Horómetro - Inicial': 'N/A',
          'Horómetro - Final': 'N/A',
          'Horómetro - OP': 'N/A',
          'Horómetro - INOP': 'N/A',
          'Horómetro - ID': 'N/A'
        });
      }
    });
    
    return data;
  }

  private prepareEstadosData(operaciones: NubeOperacion[]): any[] {
    const data: any[] = [];
    
    operaciones.forEach(op => {
      if (op.estados && op.estados.length > 0) {
        op.estados.forEach(estado => {
          // Datos base del estado
          const estadoBase = {
            'ID Operación': op.id,
            'ID Estado': estado.id,
            'Número Estado': estado.numero,
            'Estado': estado.estado,
            'Código Estado': estado.codigo,
            'Hora Inicio': estado.hora_inicio,
            'Hora Final': estado.hora_final,
            // Campos de perforación horizontal (inicialmente vacíos)
            'Perf. Horizontal - Zona': '',
            'Perf. Horizontal - Tipo Labor': '',
            'Perf. Horizontal - Labor': '',
            'Perf. Horizontal - Veta': '',
            'Perf. Horizontal - Nivel': '',
            'Perf. Horizontal - Tipo Perforación': '',
            'Perf. Horizontal - ID': '',
            // Campos de interperforación horizontal (inicialmente vacíos)
            'Ejecutado - Código Actividad': '',
            'Ejecutado - Nivel': '',
            'Ejecutado - Labor': '',
            'Ejecutado - Sección': '',
            'Ejecutado - N° Broca': '',
            'Ejecutado - N° Taladro': '',
            'Ejecutado - N° Taladros Rimados': '',
            'Ejecutado - Longitud': '',
            'Ejecutado - Detalles': '',
            'Ejecutado - ID': ''
          };

          // Procesar perforaciones horizontales si existen
          if (estado.perforaciones_horizontal && estado.perforaciones_horizontal.length > 0) {
            estado.perforaciones_horizontal.forEach(perf => {
              const rowWithPerf = {
                ...estadoBase,
                'Perf. Horizontal - Zona': perf.zona,
                'Perf. Horizontal - Tipo Labor': perf.tipo_labor,
                'Perf. Horizontal - Labor': perf.labor,
                'Perf. Horizontal - Veta': perf.veta,
                'Perf. Horizontal - Nivel': perf.nivel,
                'Perf. Horizontal - Tipo Perforación': perf.tipo_perforacion,
                'Perf. Horizontal - ID': perf.id
              };

              // Procesar interperforaciones horizontales si existen
              if (perf.inter_perforaciones_horizontal && perf.inter_perforaciones_horizontal.length > 0) {
                perf.inter_perforaciones_horizontal.forEach(inter => {
                  data.push({
                    ...rowWithPerf,
                    'Ejecutado - Código Actividad': inter.codigo_actividad,
                    'Ejecutado - Nivel': inter.nivel,
                    'Ejecutado - Labor': inter.labor,
                    'Ejecutado - Sección': inter.seccion_la_labor,
                    'Ejecutado - N° Broca': inter.nbroca,
                    'Ejecutado - N° Taladro': inter.ntaladro,
                    'Ejecutado - N° Taladros Rimados': inter.ntaladros_rimados,
                    'Ejecutado - Longitud': inter.longitud_perforacion,
                    'Ejecutado - Detalles': inter.detalles_trabajo_realizado,
                    'Ejecutado - ID': inter.id
                  });
                });
              } else {
                data.push(rowWithPerf);
              }
            });
          } else {
            data.push(estadoBase);
          }
        });
      } else {
        data.push({
          'ID Operación': op.id,
          'Mensaje': 'No hay estados registrados para esta operación'
        });
      }
    });
    
    return data;
  }

  private prepareChecklistData(operaciones: NubeOperacion[]): any[] {
    const data: any[] = [];
    
    operaciones.forEach(op => {
      if (op.checklists && op.checklists.length > 0) {
        op.checklists.forEach(check => {
          data.push({
            'ID Operación': op.id,
            'ID Checklist': check.id,
            'Descripción': check.descripcion,
            'Decisión': check.decision === 1 ? 'Aprobado' : 'Rechazado',
            'Observación': check.observacion,
            'Categoría': check.categoria
          });
        });
      } else {
        data.push({
          'ID Operación': op.id,
          'Mensaje': 'No hay checklist registrado para esta operación'
        });
      }
    });
    
    return data;
  }

  private adjustColumnWidth(worksheet: XLSX.WorkSheet, data: any[]) {
    if (!data || data.length === 0) return;

    const columnWidths: XLSX.ColInfo[] = [];
    const headers = Object.keys(data[0]);

    headers.forEach((header, i) => {
      let maxWidth = header.length * 1.2;
      
      data.forEach(row => {
        const value = row[header];
        if (value !== undefined && value !== null) {
          const length = value.toString().length;
          if (length > maxWidth) {
            maxWidth = length * 1.1;
          }
        }
      });

      const width = Math.min(Math.max(maxWidth, 10), 50);
      columnWidths.push({ wch: width });
    });

    worksheet['!cols'] = columnWidths;
  }
}