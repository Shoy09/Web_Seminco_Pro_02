import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeOperacion } from '../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelExportService {
  exportOperacionesToExcel(datosOperacionesExport: NubeOperacion[], fileName: string) {
    // Preparar datos para cada hoja
    const ejecutadoData = this.prepareEjecutadoData(datosOperacionesExport);
    const estadosData = this.prepareEstadosData(datosOperacionesExport);
    const checklistData = this.prepareChecklistData(datosOperacionesExport);

    // Crear un nuevo libro de trabajo
    const wb: XLSX.WorkBook = XLSX.utils.book_new();

    // Añadir hojas al libro con estilos de encabezado
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
    XLSX.writeFile(wb, `${fileName}_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  private prepareEjecutadoData(operaciones: NubeOperacion[]): any[] {
    const data: any[] = [];
    
    operaciones.forEach(op => {
      // Si hay horómetros, crear una fila por cada horómetro
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
        // Si no hay horómetros, exportar solo datos de operación
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
            // Campos de perforación (inicialmente vacíos)
            'Perf. - Zona': '',
            'Perf. - Tipo Labor': '',
            'Perf. - Labor': '',
            'Perf. - Veta': '',
            'Perf. - Nivel': '',
            'Perf. - Tipo Perforación': '',
            'Perf. - ID': '',
            // Campos de interperforación (inicialmente vacíos)
            'Ejecutado - Código Actividad': '',
            'Ejecutado - Nivel': '',
            'Ejecutado - Tajo': '',
            'Ejecutado - N° Broca': '',
            'Ejecutado - N° Taladro': '',
            'Ejecutado - N° Barras': '',
            'Ejecutado - Longitud': '',
            'Ejecutado - Ángulo': '',
            'Ejecutado - N° Filas': '',
            'Ejecutado - Detalles': '',
            'Ejecutado - ID': ''
          };

          // Procesar perforaciones taladro largo si existen
          if (estado.perforaciones_taladro_largo && estado.perforaciones_taladro_largo.length > 0) {
            estado.perforaciones_taladro_largo.forEach(perf => {
              const rowWithPerf = {
                ...estadoBase,
                'Perf. - Zona': perf.zona,
                'Perf. - Tipo Labor': perf.tipo_labor,
                'Perf. - Labor': perf.labor,
                'Perf. - Veta': perf.veta,
                'Perf. - Nivel': perf.nivel,
                'Perf. - Tipo Perforación': perf.tipo_perforacion,
                'Perf. - ID': perf.id
              };

              // Procesar interperforaciones si existen
              if (perf.inter_perforaciones && perf.inter_perforaciones.length > 0) {
                perf.inter_perforaciones.forEach(inter => {
                  data.push({
                    ...rowWithPerf,
                    'Ejecutado - Código Actividad': inter.codigo_actividad,
                    'Ejecutado - Nivel': inter.nivel,
                    'Ejecutado - Tajo': inter.tajo,
                    'Ejecutado - N° Broca': inter.nbroca,
                    'Ejecutado - N° Taladro': inter.ntaladro,
                    'Ejecutado - N° Barras': inter.nbarras,
                    'Ejecutado - Longitud': inter.longitud_perforacion,
                    'Ejecutado - Ángulo': inter.angulo_perforacion,
                    'Ejecutado - N° Filas': inter.nfilas_de_hasta,
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
        // Si no hay estados, exportar solo datos básicos de operación
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
        // Si no hay checklist, exportar solo datos básicos de operación
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

    // Calcular el ancho máximo para cada columna
    headers.forEach((header, i) => {
      // Ancho mínimo basado en el encabezado
      let maxWidth = header.length * 1.2;
      
      // Verificar el contenido de cada fila
      data.forEach(row => {
        const value = row[header];
        if (value !== undefined && value !== null) {
          const length = value.toString().length;
          if (length > maxWidth) {
            maxWidth = length * 1.1;
          }
        }
      });

      // Limitar entre 10 y 50 caracteres de ancho
      const width = Math.min(Math.max(maxWidth, 10), 50);
      columnWidths.push({ wch: width });
    });

    worksheet['!cols'] = columnWidths;
  }
}