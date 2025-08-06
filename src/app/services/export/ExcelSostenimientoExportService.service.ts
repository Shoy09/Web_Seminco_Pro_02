import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { NubeOperacion } from '../../models/operaciones.models';

@Injectable({
  providedIn: 'root'
})
export class ExcelSostenimientoExportService {
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
    XLSX.writeFile(wb, `${fileName}_Sostenimiento_${new Date().toISOString().slice(0,10)}.xlsx`);
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
            // Campos de sostenimiento (inicialmente vacíos)
            'Sost. - Zona': '',
            'Sost. - Tipo Labor': '',
            'Sost. - Labor': '',
            'Sost. - Veta': '',
            'Sost. - Nivel': '',
            'Sost. - Tipo Perforación': '',
            'Sost. - ID': '',
            // Campos de inter-sostenimiento (inicialmente vacíos)
            'Ejecutado - Código Actividad': '',
            'Ejecutado - Nivel': '',
            'Ejecutado - Labor': '',
            'Ejecutado - Sección': '',
            'Ejecutado - N° Broca': '',
            'Ejecutado - N° Taladro': '',
            'Ejecutado - Longitud': '',
            'Ejecutado - Malla Instalada': '',
            'Ejecutado - ID': ''
          };

          // Procesar sostenimientos si existen
          if (estado.sostenimientos && estado.sostenimientos.length > 0) {
            estado.sostenimientos.forEach(sost => {
              const rowWithSost = {
                ...estadoBase,
                'Sost. - Zona': sost.zona,
                'Sost. - Tipo Labor': sost.tipo_labor,
                'Sost. - Labor': sost.labor,
                'Sost. - Veta': sost.veta,
                'Sost. - Nivel': sost.nivel,
                'Sost. - Tipo Perforación': sost.tipo_perforacion,
                'Sost. - ID': sost.id
              };

              // Procesar inter-sostenimientos si existen
              if (sost.inter_sostenimientos && sost.inter_sostenimientos.length > 0) {
                sost.inter_sostenimientos.forEach(inter => {
                  data.push({
                    ...rowWithSost,
                    'Ejecutado - Código Actividad': inter.codigo_actividad,
                    'Ejecutado - Nivel': inter.nivel,
                    'Ejecutado - Labor': inter.labor,
                    'Ejecutado - Sección': inter.seccion_de_labor,
                    'Ejecutado - N° Broca': inter.nbroca,
                    'Ejecutado - N° Taladro': inter.ntaladro,
                    'Ejecutado - Longitud': inter.longitud_perforacion,
                    'Ejecutado - Malla Instalada': inter.malla_instalada,
                    'Ejecutado - ID': inter.id
                  });
                });
              } else {
                data.push(rowWithSost);
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